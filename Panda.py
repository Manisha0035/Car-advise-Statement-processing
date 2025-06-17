import pandas as pd

# Sample DataFrame
data = {'Name': ['Alice', 'Bob', 'Charlie'], 'Age': [25, 30, 35]}
df = pd.DataFrame(data)

# Use XlsxWriter engine to export the DataFrame to Excel
with pd.ExcelWriter('output.xlsx', engine='xlsxwriter') as writer:
    df.to_excel(writer, index=False, sheet_name='Sheet1')
