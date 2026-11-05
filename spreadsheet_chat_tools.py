"""Read-only numerical tools for the conversational spreadsheet analyst.

The model discovers ranges through Excel MCP; no report names or column mapping
are encoded here. SQL runs against temporary, in-memory copies of selected cells.
"""
import argparse
import builtins
import csv
import io
import json
from pathlib import Path
import sqlite3
import time
from uuid import uuid4

import spreadsheet_compare as compare


class AnalysisTools:
    def __init__(self, manifest_path):
        self.folder = Path(manifest_path).resolve().parent
        self.manifest = json.loads(Path(manifest_path).read_text(encoding='utf-8'))
        self.allowed = {item['id'] for item in self.manifest['files']}

    def book(self, book_id):
        if book_id not in self.allowed:
            raise ValueError('Use an attached workbook ID from the file manifest.')
        return compare.workbook(book_id)

    def record(self, item):
        with (self.folder / 'evidence.jsonl').open('a', encoding='utf-8') as stream:
            stream.write(json.dumps(item, ensure_ascii=False, allow_nan=False) + '\n')
        return item

    def cell_details(self, workbook: str, sheet: str, range: str) -> dict:
        """Inspect saved values, formula text and Excel number formats together.

        Use after Excel MCP inspection when display precision or saved formula
        results matter. Up to 2,000 cells per call. No recomputation or guessing.
        """
        book = self.book(workbook)
        data = book['sheets'].get(sheet)
        if data is None:
            raise ValueError('Unknown sheet.')
        left, top, right, bottom = compare.bounds(range)
        if (right-left+1)*(bottom-top+1) > 2000:
            raise ValueError('Inspect at most 2,000 cells; request smaller ranges.')
        cells = [{'cell': compare.address(row,col), **compare.cell(data,row,col)}
                 for row in builtins.range(top,bottom+1) for col in builtins.range(left,right+1)]
        return {'file':book['name'],'sheet':sheet,'range':range,'cells':cells,'value_source':'saved workbook cache'}

    def verify_checks(self, rules: list[dict]) -> dict:
        """Verify explicit calculation/comparison rules inferred from the actual workbook.

        Each calculation rule has kind='calculation', name, left={workbook:ID,sheet:NAME},
        operation=average|sum|min|max|count|weighted_average|ratio_of_sums,
        direction=rows|columns, source_range, target_range, optional label_range,
        number_mode=number|percent|percentage_points, tolerance (absolute units,
        or percentage points in percent modes), missing_policy=skip|error,
        ignore_zero=false, optional weights_range and rounding_digits.
        Formula source cells use their SAVED results; unavailable caches are flagged.

        Comparison rules have kind='comparison', name, left/right={workbook:ID,sheet:NAME},
        left_range/right_range (data only), matching=keys|position,
        left_keys/right_keys (comma-separated Excel column letters, composite allowed),
        left_columns/right_columns (corresponding value columns),
        value_type=auto|numeric|text, number_mode, tolerance, ignore_case=true.
        Duplicate keys are reported, never silently matched. All ranges are A1 notation.
        Returns counts, first 100 issues and first 5 other findings; complete evidence
        is saved for the user. This is a calculation executor, not a fixed report template.
        """
        if not isinstance(rules, list) or not 1 <= len(rules) <= 20:
            raise ValueError('Provide 1–20 rules.')
        for rule in rules:
            for side in ('left', 'right'):
                if side in rule:
                    self.book(rule[side].get('workbook'))
        result = compare.run(rules)
        okay = {'match', 'rounding', 'within_tolerance', 'no_data'}
        issues = [row for row in result['results'] if row['status'] not in okay]
        samples = [row for row in result['results'] if row['status'] in okay][:5]
        return self.record({'type': 'checks', 'result_id': result['id'],
                            'counts': result['counts'], 'total': result['total'],
                            'rules': result['rules'], 'issue_count': len(issues),
                            'issues': issues[:100], 'issues_truncated': len(issues) > 100,
                            'other_samples': samples})

    def query_data(self, tables: list[dict], sql: str) -> dict:
        """Run read-only SQLite SELECT for flexible analysis beyond standard checks.

        tables=[{workbook:ID,sheet:NAME,range:'B5:F30'}, ...] become t1,t2,...
        Column names retain Excel letters (B,C,D,E,F); _row is the ORIGINAL row
        number. Exclude headers/totals unless the query needs them. Cells retain
        stored values; formula cells without saved results cause an error.
        SQL supports joins, CTEs, groupings, window functions, arithmetic and filters.
        NUM(x) converts numeric and percent text to numbers, skips blank/N/A,
        rejects invalid text/errors. MEAN(x,y,...) and SUMN(x,y,...) ignore missing
        values, INCLUDE zeros, reject invalid values. SQL AVG ignores NULL.
        Always select _row/keys and reported/computed values as evidence where useful.
        Up to 500k input cells, 50k output rows, 10 seconds execution.
        Returns first 200 rows and full downloadable result artifact with source ranges
        and exact SQL. Truncation is explicit; never claim a preview is all results.
        """
        if not isinstance(tables, list) or not 1 <= len(tables) <= 10:
            raise ValueError('Provide 1–10 source tables.')
        if not isinstance(sql, str) or not 1 <= len(sql) <= 20000:
            raise ValueError('Provide a SELECT query of at most 20,000 characters.')
        db = sqlite3.connect(':memory:')
        budget, sources = 0, []
        def num(value):
            state, parsed = compare.numeric({'value': value, 'format': 'General'})
            if state == 'missing':
                return None
            if state != 'number':
                raise ValueError(f'Invalid numeric input: {value}')
            return float(parsed)
        def numbers(values):
            return [number for value in values if (number := num(value)) is not None]
        def mean(*values):
            values = numbers(values)
            return sum(values) / len(values) if values else None
        def sumn(*values):
            values = numbers(values)
            return sum(values) if values else None
        db.create_function('NUM', 1, num)
        db.create_function('MEAN', -1, mean)
        db.create_function('SUMN', -1, sumn)
        try:
            for index, spec in enumerate(tables, 1):
                book = self.book(spec.get('workbook'))
                sheet = book['sheets'].get(spec.get('sheet'))
                if sheet is None:
                    raise ValueError('Unknown sheet.')
                left, top, right, bottom = compare.bounds(spec.get('range'))
                budget += (right - left + 1) * (bottom - top + 1)
                if budget > compare.MAX_CELLS:
                    raise ValueError('Query inputs exceed 500,000 cells.')
                letters = [compare.get_column_letter(col) for col in range(left, right + 1)]
                db.execute(f'CREATE TABLE t{index} (_row INTEGER,' + ','.join(f'"{col}"' for col in letters) + ')')
                rows = []
                for row in range(top, bottom + 1):
                    values = []
                    for col in range(left, right + 1):
                        item = compare.cell(sheet, row, col)
                        if item.get('formula') and item['value'] is None:
                            raise ValueError(f'{spec["sheet"]}!{compare.address(row,col)} has no saved formula result. Use Excel MCP trace_formula to evaluate it or ask the user to recalculate/save.')
                        values.append(item['value'])
                    rows.append([row] + values)
                db.executemany(f'INSERT INTO t{index} VALUES (' + ','.join('?' for _ in range(len(letters)+1)) + ')', rows)
                sources.append({'table': f't{index}', 'file': book['name'], **spec})
            allowed = {sqlite3.SQLITE_SELECT, sqlite3.SQLITE_READ, sqlite3.SQLITE_FUNCTION, sqlite3.SQLITE_RECURSIVE}
            def authorize(action, first, second, *_):
                if action not in allowed or (action == sqlite3.SQLITE_FUNCTION and str(second).lower() in {'load_extension', 'readfile', 'writefile'}):
                    return sqlite3.SQLITE_DENY
                return sqlite3.SQLITE_OK
            db.set_authorizer(authorize)
            started = time.monotonic()
            db.set_progress_handler(lambda: int(time.monotonic() - started > 10), 10000)
            cursor = db.execute(sql)
            if cursor.description is None:
                raise ValueError('Only SELECT queries are supported.')
            columns = [column[0] for column in cursor.description]
            rows = cursor.fetchmany(compare.MAX_RESULTS + 1)
            if len(rows) > compare.MAX_RESULTS:
                raise ValueError('Output exceeds 50,000 rows. Narrow the query; no partial result was saved.')
            artifact_id = uuid4().hex
            artifact = {'id': artifact_id, 'columns': columns, 'rows': [list(row) for row in rows], 'sql': sql, 'sources': sources}
            compare.write_json(self.folder / 'artifacts' / (artifact_id + '.json'), artifact)
            return self.record({'type': 'query', 'artifact_id': artifact_id, 'columns': columns,
                                'rows': artifact['rows'][:200], 'total_rows': len(rows),
                                'truncated': len(rows) > 200, 'sql': sql, 'sources': sources})
        except sqlite3.Error as error:
            raise ValueError(f'Query failed: {error}. No result was saved. Check numeric inputs and use read-only SQL.') from error
        finally:
            db.close()


def artifact_csv(artifact):
    output = io.StringIO(newline='')
    writer = csv.writer(output)
    def safe(value):
        # Prevent exported strings from becoming executable spreadsheet formulas.
        return "'" + value if isinstance(value, str) and value.lstrip().startswith(('=', '+', '-', '@')) else value
    writer.writerow([safe(value) for value in artifact['columns']])
    writer.writerows([[safe(value) for value in row] for row in artifact['rows']])
    return '\ufeff' + output.getvalue()


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('manifest')
    args = parser.parse_args()
    from mcp.server.mcpserver import MCPServer
    from mcp.types import ToolAnnotations
    analyst = AnalysisTools(args.manifest)
    server = MCPServer('spreadsheet-analysis', log_level='WARNING')
    read_only = ToolAnnotations(readOnlyHint=True, destructiveHint=False,
                                idempotentHint=True, openWorldHint=False)
    server.tool(annotations=read_only)(analyst.verify_checks)
    server.tool(annotations=read_only)(analyst.query_data)
    server.tool(annotations=read_only)(analyst.cell_details)
    server.run(transport='stdio')
