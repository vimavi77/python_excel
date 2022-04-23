#codigo creado 29-nov-2020
import xlrd

excel_files="/Users/macbook/Desktop/reporte.xlsx"

workbook=xlrd.open_workbook(excel_files)

worksheet=workbook.sheet_by_index(0)

headings=worksheet.row_values(0)
#print(headings)
col2heading=headings[3]
#print(col2heading)

i =0
text="TECATE"
for row in range(worksheet.nrows):
    if str(worksheet.cell(row,1).value) == text:  #Find within Column 8 with Header Domain that matches a text
        i = i +1
        print(worksheet.cell(row,4).value)
    else:
    	#print(worksheet.cell(row,7).value)
        pass

#print("Found %s" %i)
print("Found %s Email Names that matches the %s domain" %(i,text))