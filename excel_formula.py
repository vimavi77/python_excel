#codigo creado 23-nov-2020
import openpyxl
excel_files=['/Users/macbook/Desktop/remove_duplicates.xlsx']



for file in excel_files:
	workbook=openpyxl.load_workbook(file)
	worksheet=workbook['Master']
	worksheet['L6']='=sum(L2:L5)'
	workbook.save(file)