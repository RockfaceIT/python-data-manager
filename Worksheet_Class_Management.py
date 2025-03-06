## load class
from my_worksheet import MyWorksheet

## Set Input & Output files
input_file = "input_file.xlsx"
output_file = "output_file.xlsx"

## Define your Column Headings per JotForm Form designation
column_headings = ["ID",
                   "Provider's Name",
                   "California Acupuncture License Number",
                   "Personal Email"]

## Designate the columns with corresponding number of columns to remove
columns_delete = {
    1: 1,
    3: 3,
    5: 1
}

## Create worksheet object
sheet = MyWorksheet(input_file, output_file, column_headings, columns_delete)

## Set fullname column: merge first and last name columns
sheet.set_fullname(3, 4)

## Manage Columns: delete unused columns & set first row column headings
sheet.manage_columns()

## Save File
sheet.save_file()
