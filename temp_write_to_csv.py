import openpyxl, os
import pandas as pd


temp_csv_dir = r"C:\Github local repos\create_project_summary\private\temp_csv"

csv_file = 'SurveyResults.csv'

csv_name_full = os.path.join(temp_csv_dir, csv_file)
print(csv_name_full)

"""
wb = openpyxl.load_workbook(csv_name_full)
sheet = wb.active
print(sheet['A1'])
"""
df = pd.read_csv(csv_name_full)

# print(help(df))

df['client_name'] = 'SSI'
# print(df.head())
# print(df)
new_csv_name = 'SurveyResults_w_client_name.csv'
new_csv_full = os.path.join(temp_csv_dir, new_csv_name)
df.to_csv(new_csv_full, index=None)
 # TODO: install jupyter to make working with dfs more interesting
 # TODO: move this to module function to add col to csv
 
