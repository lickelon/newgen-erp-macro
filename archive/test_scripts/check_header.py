import sys
import pandas as pd

sys.stdout.reconfigure(encoding='utf-8')

df = pd.read_excel('연말정산.xls', header=None)

print('처음 3행 (헤더 확인):\n')
for i in range(3):
    print(f'Row {i}:')
    print(list(df.iloc[i]))
    print()
