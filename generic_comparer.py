import pandas as pd
import numpy as np
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from os import listdir, path

new_excel_file = 'C:/USERDATA/General Tasks/GE Tags Analysis/TAGL/New/Tag Register_New format_26.02.2021.xlsx'
old_excel_file = 'C:/USERDATA/General Tasks/GE Tags Analysis/TAGL/Old/Tag Register_New format_01022021_BH_Feedback_1_JFJV reply.xlsx'
# new_folder_files = listdir(new_lft_folder)
file_path = 'C:/USERDATA/General Tasks/GE Tags Analysis/TAGL/New/Tag Register_New format_26.02.2021.xlsx'
# def start_compare(file_path):
print('processing ' + file_path)
filename = path.splitext(file_path)[0]  # seperating filename and extension

df_new = pd.read_excel(new_excel_file, sheet_name="Sheet2")
df_old = pd.read_excel(old_excel_file, sheet_name="Sheet2")

# getting new tags
df_new_tags = pd.concat([df_new, df_old]).drop_duplicates(subset=['Tag No'], keep=False)
new_tags_list = df_new_tags.index.values.tolist()  # holding new tag numbers
df_new = df_new.drop(new_tags_list)

df_new = df_new[df_new['Tag No'].notna()]  # removing blank rows
df_old = df_old[df_old['Tag No'].notna()]

df_new  = df_new.reset_index(drop=True)
df_old = df_old.reset_index(drop=True)

# getting the difference between two
df_new['StatusMatch'] = np.where(df_new['Status'] == df_old['Status'],'True','False')
df_new['DescriptionMatch'] = np.where(df_new['Description'] == df_old['Description'],'True','False')
difference = df_new[df_new.values != df_old.values]
# getting row and col index for differences
rowCol = np.argwhere(difference.notnull().values).tolist()

if rowCol:
    new_wb = load_workbook(new_excel_file)
    new_ws = new_wb.worksheets[0]

for row in rowCol:
    colLetter = get_column_letter(row[1]+1)
    new_ws[colLetter + str(row[0]+3)].fill = PatternFill(start_color="ededa8", end_color="ededa8",
                                                         fill_type="solid")
if rowCol:
    new_wb.save('./' + filename + '_highlighted.xlsx')

print(filename + ' processing completed')
