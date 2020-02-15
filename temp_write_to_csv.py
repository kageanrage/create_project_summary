import openpyxl, os
import pandas as pd

temp_csv_dir = r"C:\Github local repos\create_project_summary\private\temp_csv"
csv_file = 'SurveyResults.csv'
csv_name_full = os.path.join(temp_csv_dir, csv_file)

df = pd.read_csv(csv_name_full)
df['client_name'] = 'SSI'
df.to_csv(csv_name_full, index=None)
