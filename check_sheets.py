import pandas as pd

excel_file = 'data/ПОЛОТНО - 4аКШО-тексерілді.xlsx'

xl = pd.ExcelFile(excel_file)
print("Листы в файле:")
for i, sheet in enumerate(xl.sheet_names, 1):
    print(f"{i}. {sheet}")
