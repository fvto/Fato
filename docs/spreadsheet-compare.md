# Spreadsheet Chat and Compare

Open `/compare`, or choose **Spreadsheet Compare** from the shared menu, SFC Dashboard or report tools page.

## Workflow

The default mode is **So sánh ngoại tuyến · trên máy của bạn**. Choose
**ChatGPT trên web** for a manual handoff. The Dashboard prepares a
prompt locally and offers original-workbook downloads. Copy the prompt, open
ChatGPT, attach the files and send there. No Codex inference or API call runs in
this mode. ChatGPT plan/feature limits still apply. Files and prompts are not
automatically uploaded, and ChatGPT answers do not automatically sync back.

The prepared prompt asks ChatGPT for `dashboard-result.json`. Import that file or
paste its JSON under **Đưa kết quả ChatGPT về Dashboard** to display paired cards
and save the imported result in the local conversation. Imported results are
explicitly labeled as supplied by ChatGPT and not independently verified by the
Dashboard. No inference is used for import. Prior results can optionally be
included in the copied prompt when the attached workbook IDs match.

Choose **Codex · phân tích ngay trên Dashboard** or **Antigravity CLI · phân tích ngay trên Dashboard** to use the automated workflow:

1. Upload Excel workbooks (`.xlsx`, `.xlsm`, `.xls`) or import a saved Dashboard export.
2. Type a natural-language request and click **Send request**. English and Vietnamese are supported.
3. AI inspects the actual sheets/headers, chooses ranges and performs numerical checks through tools. No report-to-column mapping or keyword-to-template router is used. Ambiguous methods can be clarified in the conversation.
4. Ask follow-up questions, inspect verified cells and download evidence. Manual rules remain available under **Advanced · Set rules manually**.

Comparison findings are grouped by section and named metric/question. Each card
places the original file names, corresponding sections and values side by side,
then explains the difference. Exact sheet/cell references are collapsed under
the evidence control. Checks within one workbook name the summary/detail or
reported/recalculated sides explicitly. Cards survive reloads and remain part of
follow-up context. Existing text-only replies retain their original formatting;
new analyses use the structured layout.

## AI integration and setup

Research compared [Excel MCP Server](https://github.com/haris-musa/excel-mcp-server),
[PandasAI](https://github.com/sinaptik-ai/pandas-ai) and [mcp-excel](https://github.com/jwadow/mcp-excel).
This application integrates the MIT-licensed Excel MCP Server 2.2.2 for real workbook
inspection, with the already authenticated local Codex CLI for natural-language
reasoning and tool calling. Excel MCP itself is a tool server, not a language model.
[Codex non-interactive documentation](https://learn.chatgpt.com/docs/non-interactive-mode)
describes the ephemeral/structured-output execution used here.

Run `python scripts/setup_excel_ai.py` once to create an isolated `.runtime/excel-ai`
environment from pinned `requirements-excel-ai.txt`. Install/sign in to Codex CLI
separately. `SPREADSHEET_CODEX_CLI` can point to its executable. No API key is stored
in the Dashboard and no global Codex configuration is changed. Chat uses the signed-in
account's usage limits; AI availability depends on that connection.

### Antigravity CLI

Install [Antigravity CLI](https://antigravity.google/docs/cli/install/) and open
`agy` once to sign in. Select **Antigravity CLI** under **Cách phân tích** on
`/compare`. The Dashboard checks the signed-in account using `/usage`, without
running inference. `SPREADSHEET_ANTIGRAVITY_CLI` can point to `agy.exe` when it is
not on PATH. The existing Excel MCP runtime is shared with Codex.

Antigravity runs through the documented [headless interface](https://antigravity.google/docs/cli/headless/)
with streamed input and schema-constrained JSON results. Related workbook content
is sent through the signed-in Antigravity account and uses its quota. Extra G1
credits are disabled for Dashboard sessions. Each request records its provider;
switching providers keeps the saved conversation and evidence available.

Per-job MCP configuration and a dedicated analyst with only the built-in completion tool
are prepared in a temporary workspace outside the report repository. Session
permissions allow the two spreadsheet MCP servers and deny shell commands, file
reads/writes, and web access through built-in tools. No global MCP or permission
settings are edited. Antigravity can still load the user's personal CLI
customizations; unlike Codex, its CLI has no `--ignore-user-config` flag.
The eight spreadsheet tools are loaded eagerly so their parameter schemas are
available without granting the analyst access to local configuration files.
Original workbooks remain unchanged. Results and independently computed numerical
evidence use the same Dashboard cards and downloads as Codex.
The integration was tested with Antigravity CLI 1.5.3 using a synthetic workbook.

Per turn, Codex receives only the two configured MCP servers: upstream Excel MCP in
read-only mode confined to staged copies of the attached files, and local analysis
tools. Shell, browser, web search, plugins and additional agents are disabled.
Analysis tools expose exact rule verification, saved cell/formula/format inspection,
and bounded read-only SQLite queries over any selected ranges. Queries can aggregate,
filter and join arbitrary columns; there is no arbitrary Python execution. SQL
writes, attachments and extension loading are denied. Numerical results are saved
independently of the model's answer, with source ranges and the executed query/rules.

Relevant workbook content is sent to AI through Codex. Original files remain
unchanged on disk. Macro code is never run. Legacy `.xls` is staged as values-only
`.xlsx` for Excel MCP; its formula text is unavailable. Excel MCP may evaluate
supported formulas while the local verification engine uses saved cached values;
`cell_details` makes saved values and number formats explicit.

Jobs run in the background with a four-minute timeout and a two-job concurrency cap.
Reload restores an existing job/chat without submitting another inference. Send
requests have persistent idempotency IDs. Server interruptions are surfaced as errors.
Saved conversations can be opened with `/compare?chat=<conversation-id>`; attached
files are restored without reanalysis. Chat data and full evidence live under
`Report/Comparisons/chat` and are excluded from Git.

UI verification used a pair of workbooks with unrelated sheet names, reordered
columns and reordered records: the AI found an incorrect mean including zero and
a changed value by its record key. A real MQAA request, with no supplied ranges,
discovered `SUMMARY Shoe`, checked 51 row means and located P15/P29; Grand Total
was explicitly excluded from that request. These are integration examples, not
hardcoded defaults.

Original workbooks are copied into `Report/Comparisons/inputs` and are never changed. Results are retained in `Report/Comparisons/results`. Browser storage holds workbook IDs, rules, presets and the last result ID. Reloading restores them without submitting another run.

## Configurable calculations

- Input ranges and expected-result ranges use Excel A1 notation. The source may be any sheet; no MQAA sheet names or row numbers are hardcoded.
- Check each input row against one expected-value column, or each input column against one expected-value row.
- Supported operations: average, sum, min, max, numeric count, weighted average and ratio of sums.
- Weighted average is `sum(value × weight) / sum(weight)`. Ratio of sums is `sum(numerator) / sum(denominator)`. The second range has the same dimensions as the input range.
- Blank cells and text markers `N/A`, `NA`, `N.A.`, `-` and `—` can be skipped or reported. Zero is included unless explicitly excluded. Unexpected text, Excel errors and uncached formulas are reported rather than discarded.
- Suggested rules use visible headers and ranges. Footer suggestions deliberately require review: an average of department averages, an average of all observations and a ratio of total scores are different mathematical rules.

## Numbers and rounding

**Numbers as stored** preserves numeric magnitudes. Text `94.4%` is interpreted as `0.944`.

**Percent strict** accepts fractional numeric values and percent text / formatted cells. Unformatted `30` is flagged rather than guessed to mean 30%.

**Percent points** interprets unformatted numeric `30` as 30%, while a percent-formatted numeric `0.3` remains 30%.

Percent tolerances are entered in percentage points: `0.05` permits a difference of `0.0005` in fractional storage. Ordinary number tolerances use absolute units. Display decimals can be supplied or inferred from percentage formats / percentage text. Rounding matches and values within tolerance have separate statuses.

## Dataset comparisons

Choose files, sheets, data ranges excluding headers, corresponding value columns, and keys such as `A,B` on each side. A key's column may differ between the two datasets. Rows may be reordered. Duplicate or invalid keys are reported with original row numbers; no first-match or many-to-many join is performed. Position matching is available when explicitly selected. Numeric, text and automatic value comparisons are supported.

## Verification limits

The manual rule checker uses deterministic calculations and the workbook's saved formula cache. Uncached formulas are reported. Chat can inspect/evaluate supported formulas through Excel MCP, but unsupported formulas still require recalculation and saving in Excel / WPS. Macros and external links are never executed by local analysis tools. `.xls` supplies saved values through xlrd rather than formula text. Chart/image interpretation is outside this integration's scope. AI can misunderstand a business method: review its stated scope and evidence, especially for weighted totals.

Each workbook is limited to 40 MB and 500,000 scanned cells. A run allows up to 20 rules, 1,000,000 selected cells and 50,000 findings. The UI pages 200 findings at a time; downloads contain the complete result. Oversized runs are rejected rather than silently truncated.

## Tests

`python -m unittest tests.test_spreadsheet_compare tests.test_spreadsheet_chat -q`

`node tests/test_spreadsheet_chat_format.js`

`node tests/test_spreadsheet_chatgpt.js`

Optional Chrome check: set `CHINGLUH_UI_TESTS=1` and run
`python -m unittest tests.test_spreadsheet_provider_browser -q`.

Coverage includes zeros, missing values, percentages, rounding, uncached formulas, row/column aggregates, weights, composite keys, reordered rows, duplicate keys, upload validation, result pagination, local request boundaries and downloads. Browser verification uses the referenced MQAA workbook: 51 row-average checks and 13 column-total checks locate P15, P29 and J56 under the chosen mean/tolerance rules.
