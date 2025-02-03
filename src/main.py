from openpyxl import load_workbook
import datetime as dt
import os
import xlsxwriter
import json
import asyncio
from pyppeteer import launch
import platform
from xlsx2html import xlsx2html
import sys
from tkinter import messagebox

std_out = sys.stdout
file_out = open("log.txt", "w")
sys.stdout = file_out
def HandleError(e):
	messagebox.showerror("Error", f"An error occurred: {str(e)}")
	print(e)
	sys.stdout = std_out
	file_out.close()
	sys.exit(1)
	
async def generate_pdf(url, pdf_path):
	try:
		bpth = os.getcwd()+'/chrome-win/chrome.exe'
		print(bpth)
		if platform.system() != "Windows":
			bpth = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
		browser = await launch(headless=True, executablePath=bpth, options={'args': ['--no-sandbox']})
		page = await browser.newPage()
		
		await page.goto(url)
		
		await page.pdf({'path': pdf_path, 'format': 'A4', 'printBackground': True})
	except Exception as e:
		print(f"Error generating PDF: {e}")
	finally: 	
		await browser.close()
	# converter.convert(f'{url}', f'{pdf_path}')

def to_pdf(file_path, in_root):
	print(f"Generating PDF for of {file_path}...")
	file_path = file_path.replace('\\', '/')
	os.path.relpath(file_path, os.getcwd()) 
	print(file_path)
	xlsx2html(file_path, 'file.html')
	asyncio.run(generate_pdf(f'file:///{os.getcwd()+'/file.html'}', file_path.replace('.xlsx', '.pdf')))

def handle_excel_write(file_path, personal_data):
	# Create a workbook and add a worksheet.
	workbook = xlsxwriter.Workbook(file_path)
	worksheet = workbook.add_worksheet()
	# Add a bold format to use to highlight cells.
	bold = workbook.add_format({'bold': True})
	# add a blue background to the headers.
	header_format = workbook.add_format({'bold': True, 'bg_color': '#95b3d7', 'align': 'center', 'valign': 'vcenter', 'border': 1})
	# add border to the cells.
	cell_format = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})
	worksheet.write('A1', 'UAN NUMBER', header_format)
	worksheet.write('B1', 'MEMBER ID', header_format)
	worksheet.write('C1', 'ESTABLISHMENT DETAILS', header_format)
	worksheet.write('D1', 'NAME', header_format)
	worksheet.write('E1', 'FATHER OR HUSBAND \n NAME', header_format)
	worksheet.write('F1', 'DATE OF JOIN', header_format)
	worksheet.write('G1', 'DATE OF EXIT PF', header_format)
	last_row = 1
	# Start from the first cell below the headers.
	for row, index in enumerate(personal_data): # has the following format: [str(uan), row['MemID'], est, Name, FName, dt.datetime.strptime(row["EPFDOJ"], '%d-%b-%Y'), ""]
		for col, item in enumerate(index):
			if col >= 5:
				if type(item) == str:
					worksheet.write(row + 1, col, item, cell_format)
				else:
					worksheet.write(row + 1, col, item.strftime('%Y-%m-%d'), cell_format)
			else:
				worksheet.write(row + 1, col, item, cell_format)
		last_row += 1
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
			HandleError(f"No exit date for UAN {uan}")
			FName = ""
			Name = name

		# est = company_names_map.get(row["EstID"], row["EstID"])
		try:
			est = company_names_map[row["EstID"]]
		except KeyError:
			HandleError(f"No company name for EstID {row['EstID']}")
			est = row["EstID"]
		# print(output)

		if row["EPFDOE"].strip() == "":
			output.append([str(uan), row['MemID'], est, Name, FName, dt.datetime.strptime(row["EPFDOJ"], '%Y-%m-%d'), ""])
			
		else:
			output.append([str(uan), row['MemID'], est, Name, FName, dt.datetime.strptime(row["EPFDOJ"], '%Y-%m-%d'), dt.datetime.strptime(row["EPFDOE"], '%Y-%m-%d')])
	output = sorted(output, key=lambda x: x[5])[::-1]
	handle_excel_write(os.path.join(root, f"{uan}.xlsx"), output)
	workbook.close()
	return output

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

	# print(json.dumps(company_names_map, indent=2))
	company_name_workbook.close()


	uan_id_workbook = load_workbook(filename=uan_id_compiled_file)
	uan_id_sheet = uan_id_workbook.active
	uan_id_map = []
	for row in uan_id_sheet.iter_rows(min_row=2,
							min_col=0,
							max_col=1,
							values_only=True):
		uan_id_map.append(row[0])
	# print(json.dumps(uan_id_map, indent=2))
	uan_id_workbook.close()

	uan_name_workbook = load_workbook(filename=uan_name_compiled_file)
	uan_name_sheet = uan_name_workbook.active
	uan_name_map = {}
	for row in uan_name_sheet.iter_rows(min_row=0,
							min_col=0,
							max_col=3,
							values_only=True):
		print(row)
		uan_name_map.update({str(row[0]): {'Name': row[1], 'FName': row[2]}})
	# print(json.dumps(uan_name_map, indent=2))
	uan_name_workbook.close()


	data = []
	for root, dirs, files in os.walk(in_root):
		path = root.split(os.sep)
		for file in files:
			if "uanacord_uan_member_data" in file:
				if file[-17:-5] in uan_name_map.keys():
					print(f"Processing {file}...")
				else:
					print(f"UAN ID not found for {file}...")
				data.append(handle_input(os.path.join(root, file), file[-17:-5], out_root, uan_name_map, company_names_map))
	
	# convert all xlsx files to pdf
	for root, dirs, files in os.walk(out_root):
		for file in files:
			if file.endswith(".xlsx"):
				to_pdf(os.path.join(root, file), in_root)
	
	combine_xlsx(out_root, data)

def combine_xlsx(out_root, data):
	print("Combining xlsx files...")
	# reads all the xlsx files in the directory and combines all the data into one xlsx file and saves it, uses xlsx writer
	# seperate tables for each uan id
	# Create a workbook and add a worksheet.
	workbook = xlsxwriter.Workbook(os.path.join(out_root, "Excel Reports.xlsx"))
	worksheet = workbook.add_worksheet()
	# header's format.
	header_format = workbook.add_format({'bold': True, 'bg_color': '#95b3d7','align': 'center', 'valign': 'vcenter', 'border': 1})
	# add border to the cells.
	cell_format = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})
	row = 0
	col = 0
	for group in data:
		
		worksheet.write(row, 0, 'UAN NUMBER', header_format)
		worksheet.write(row, 1, 'MEMBER ID', header_format)
		worksheet.write(row, 2, 'ESTABLISHMENT DETAILS', header_format)
		worksheet.write(row, 3, 'NAME', header_format)
		worksheet.write(row, 4, 'FATHER OR HUSBAND \n NAME', header_format)
		worksheet.write(row, 5, 'DATE OF JOIN', header_format)
		worksheet.write(row, 6, 'DATE OF EXIT PF', header_format)
		for r in group:
			for col, value in enumerate(r):
				worksheet.write(row + 1, col, value, cell_format)
			row += 1
			col = 0
	worksheet.autofit()

	workbook.close()
	os.remove("file.html")
	print("Excel Reports.xlsx has been created.")

if __name__ == "__main__":
	company_name_map_file = "C:/Users/ndben/Desktop/Nikhil/Freelancing/IVerify/src/Tests/Data Base for test new deleted.xlsx"
	uan_id_compiled_file = "C:/Users/ndben/Desktop/Nikhil/Freelancing/IVerify/src/Tests/Uan and candidate.xlsx"
	uan_name_compiled_file = "C:/Users/ndben/Desktop/Nikhil/Freelancing/IVerify/src/Tests/name and father name.xlsx"
	in_root = "C:/Users/ndben/Desktop/Nikhil/Freelancing/IVerify/src/Tests/Website Data/Website Data"
	out_root = "C:/Users/ndben/Desktop/Nikhil/Freelancing/IVerify/src/Tests/Website Data/Out"
	# company_name_map_file = input("Enter the path to the Company name map file: ")
	# uan_id_compiled_file = input("Enter the path to the UAN ID compiled file: ")
	# uan_name_compiled_file = input("Enter the path to the UAN Name compiled file: ")
	# in_root = input("Enter the base directory path: ")
	# out_root = input("Enter the out directory path: ")
	compute(company_name_map_file, uan_id_compiled_file, uan_name_compiled_file, in_root, out_root)
	
	print("Completed.")