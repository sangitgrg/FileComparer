# FileComparer

This program uses pandas and numpy to compare two big excel files and create output with highlight in each cell that are not matched.
gener_compare.py is the main script for comparing.
Ignore other files, they are under development.

# config ini file structure

[PathConfig]
newFolderPath = D:/USERDATA/NewFolder # keep excel file with new data
oldFolderPath = D:/USERDATA/OldFolder # keep excel file with old data

[ExcelConfig]
newExcelsheetName='Sheet1' # sheet name for your excel file
oldExcelsheetName='Sheet1'
uniqueKey='Tag No' # this unique key will be used as a base for comparing between two excel files

# Limitation

1. Currently only excel file is supported.
1. Comaprison should be done by taking one unique key/column from excel as a base for comparison.
1. Row and column count should match between new and old excel file.
