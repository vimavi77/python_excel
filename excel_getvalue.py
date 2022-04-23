#codigo creado 24-abr-2021
import openpyxl

excel_files=['/Users/macbook/Desktop/remove_duplicates.xlsx']

values=[]
import warnings
warnings.simplefilter("ignore")
for file in excel_files:
	workbook=openpyxl.load_workbook(file) #Open Workbook file
	worksheet=workbook['Master'] #Select Worksheet
	cell_value=worksheet['k4'].value

print(cell_value)