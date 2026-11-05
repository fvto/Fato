import pandas as pd

vh_file = r'Report\ToolRuns\4298bb50bf9e4d55b1bf8b214482c0cf\workspace\HFPA\HFPA_VH_Internal_2026-10-01_2026-10-31-VH.xlsx'
df = pd.read_excel(vh_file)

# filter models containing FRANCHISE
sub = df[df['Model'].str.contains('FRANCHISE', case=False, na=False)]
print("Unique franchise models:", sub['Model'].unique())
def_cols = [c for c in df.columns[21:] if sub[c].sum() > 0]
total_def = sum(sub[c].sum() for c in def_cols)
print(f"Total defects: {total_def}")
for c in def_cols:
    val = sub[c].sum()
    pct = (val / total_def) * 100 if total_def else 0
    if pct > 4.0:
        print(f"  {c}: {val} ({pct:.2f}%)")
