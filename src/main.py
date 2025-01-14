import json
from openpyxl import load_workbook

workbook = load_workbook(filename=input("Enter the file name/path: "))
sheet = workbook.active

products = {}

# Using the values_only because you want to return the cells' values
for row in sheet.iter_rows(min_row=2,
                           min_col=0,
                           max_col=7,
                           values_only=True):
    print(row)

# Using json here to be able to format the output for displaying later
print(json.dumps(products))
