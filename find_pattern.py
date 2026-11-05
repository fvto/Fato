import os
import openpyxl

patterns = ['13.10%', '0.1310', '0.1236', '12.36%', '9.83%']

def check_excel(path):
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        for s in wb.sheetnames:
            ws = wb[s]
            for row in ws.iter_rows(max_row=60, max_col=30, values_only=True):
                r_str = ' '.join(str(c) for c in row if c is not None)
                for pat in patterns:
                    if pat in r_str:
                        print(f"FOUND {pat} in Excel: {path} [sheet {s}]")
        wb.close()
    except Exception as e:
        pass

for root, dirs, files in os.walk('.'):
    if any(skip in root for skip in ['.git', 'node_modules', '.gemini', 'Lib', 'Scripts', 'venv']):
        continue
    for f in files:
        if f.endswith('.xlsx') and not f.startswith('~$'):
            check_excel(os.path.join(root, f))
print("Search finished.")

