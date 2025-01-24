import json
from openpyxl import load_workbook
import datetime as dt
import os
import xlsxwriter
import xlwings as xw
def to_pdf(file_path):
	excel_app = xw.App(visible=False)
	excel_book = excel_app.books.open(file_path)
	excel_book.sheets[0].to_pdf(path=file_path.replace('.xlsx', '.pdf'), layout=None, show=False, quality='standard')
	excel_book.save()
	excel_book.close()
	excel_app.quit()
	# # Open Microsoft Excel 
	# excel = client.Dispatch("Excel.Application") 
	
	# # Read Excel File 
	# sheets = excel.Workbooks.Open(file_path) 
	# work_sheets = sheets.Worksheets[0] 
	
	# # Convert into PDF File 
	# work_sheets.ExportAsFixedFormat(0, file_path.replace('.xlsx', '.pdf')) 

def handle_excel_write(file_path, personal_data):

	# Create a workbook and add a worksheet.
	workbook = xlsxwriter.Workbook(file_path)
	worksheet = workbook.add_worksheet()

	last_row = 1
	# Start from the first cell below the headers.
	for row, index in enumerate(personal_data): # has the following format: [str(uan), row['MemID'], est, Name, FName, dt.datetime.strptime(row["EPFDOJ"], '%d-%b-%Y'), ""]
		for col, item in enumerate(index):
			if col >= 5:
				if type(item) == str:
					worksheet.write(row + 1, col, item)
				else:
					worksheet.write(row + 1, col, item.strftime('%d-%b-%Y'))
			else:
				worksheet.write(row + 1, col, item)
		last_row += 1

	worksheet.add_table(f'A1:G{last_row}', {'autofilter': False, 'banded_columns': False, 'banded_rows': False, 
											'style': 'Table Style Medium 2',
											'columns': [{'header': 'UAN NUMBER'},
														{'header': 'MEMBER ID'},
														{'header': 'ESTABLISHMENT DETAILS'},
														{'header': 'NAME'},
														{'header': 'FATHER OR HUSBAND \n NAME'},
														{'header': 'DATE OF JOIN'},
														{'header': 'DATE OF EXIT PF'},
														]}, )
	worksheet.autofit()

	workbook.close()
	


def handle_input(file_path, uan, root, uan_name_map, company_names_map):
	output = []
	workbook = load_workbook(filename=file_path)
	sheet = workbook.active

	personal_data = {}

	# Using the values_only because you want to return the cells' values
	for row in sheet.iter_rows(min_row=4,
							min_col=0,
							max_col=14,
							values_only=True):
		# print('----', file_path, '----')
		# print(row, len(row))
		personal_data.update({row[0]: {'MemID': row[0], 
									   'EstID': row[1],
									   'EPFDOJ': row[2], 
									   'EPSDOJ': row[3],
									   'EPFDOE': row[4],
									   'EPSDOE': row[5],
									   'ROL': row[6],
									   'PE': row[7], 
									   'BT': row[8],
									   'CS': row[9],
									   'BalSt': row[10],
									   'NEMem': row[11],
									   'UANLink': row[12],}})

	for row in personal_data.values():
		name = uan_name_map.get(uan, uan)
		if type(name) == dict:
			FName = name.get("FName", "").upper()
			Name = name.get("Name", "").upper()
		else:
			FName = ""
			Name = name

		est = company_names_map.get(row["EstID"], row["EstID"])
		
		if row["EPFDOE"].strip() == "":
			output.append([str(uan), row['MemID'], est, Name, FName, dt.datetime.strptime(row["EPFDOJ"], '%d-%b-%Y'), ""])
		else:
			output.append([str(uan), row['MemID'], est, Name, FName, dt.datetime.strptime(row["EPFDOJ"], '%d-%b-%Y'), dt.datetime.strptime(row["EPFDOE"], '%d-%b-%Y')])
	output = sorted(output, key=lambda x: x[5])[::-1]
	print("___________")
	handle_excel_write(os.path.join(root, f"{uan}.xlsx"), output)
	print("___________")
	workbook.close()

def compute(company_name_map_file, uan_id_compiled_file, uan_name_compiled_file, in_root, out_root):
	if not os.path.exists(out_root):
		os.makedirs(out_root)	
	company_name_workbook = load_workbook(filename=company_name_map_file)
	company_name_sheet = company_name_workbook.active
	company_names_map = {}

	# Using the values_only because you want to return the cells' values
	for row in company_name_sheet.iter_rows(min_row=2,
							min_col=0,
							max_col=2,
							values_only=True):
		company_names_map.update({row[0]:  row[1]})

	# print(json.dumps(company_names_map))
	company_name_workbook.close()


	uan_id_workbook = load_workbook(filename=uan_id_compiled_file)
	uan_id_sheet = uan_id_workbook.active
	uan_id_map = []
	for row in uan_id_sheet.iter_rows(min_row=2,
							min_col=0,
							max_col=1,
							values_only=True):
		# print("row")
		uan_id_map.append(row[0])
	# print(json.dumps(uan_id_map))
	uan_id_workbook.close()

	uan_name_workbook = load_workbook(filename=uan_name_compiled_file)
	uan_name_sheet = uan_name_workbook.active
	uan_name_map = {}
	for row in uan_name_sheet.iter_rows(min_row=0,
							min_col=2,
							max_col=4,
							values_only=True):
		uan_name_map.update({row[0]: {'Name': row[1], 'FName': row[2]}})
	uan_name_workbook.close()


	for root, dirs, files in os.walk(in_root):
		path = root.split(os.sep)
		for file in files:
			if "uanacord_uan_member_data" in file:
				if file[-17:-5] in uan_name_map.keys():
					print(f"Processing {file}...")
				else:
					print(f"UAN ID not found for {file}...")
				handle_input(os.path.join(root, file), file[-17:-5], out_root, uan_name_map, company_names_map)
	for root, dirs, files in os.walk(out_root):
		for file in files:
			if file.endswith(".xlsx"):
				to_pdf(os.path.join(root, file))
		

if __name__ == "__main__":
	company_name_map_file = "sample/CompDat.xlsx" # input("Enter the path to the Company name map file: ")
	uan_id_compiled_file = "sample/UAN.xlsx" # input("Enter the path to the UAN ID compiled file: ")
	uan_name_compiled_file = "sample/UANFName.xlsx" # input("Enter the path to the UAN Name compiled file: ")
	in_root = "sample/in" # input("Enter the base directory path: ")
	out_root = "sample/outers" # input("Enter the out directory path: ")
	compute(company_name_map_file, uan_id_compiled_file, uan_name_compiled_file, in_root, out_root)