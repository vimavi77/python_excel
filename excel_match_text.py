#codigo creado 18-may-2021
import xlrd

excel_files="/Users/macbook/Desktop/remove_duplicates.xlsx"

workbook=xlrd.open_workbook(excel_files)

worksheet=workbook.sheet_by_index(0)

headings=worksheet.row_values(0) # Select the headers of the row=1
#print(headings)
col2heading=headings[3]   # Select the 4th header  wich is column D and the row=1 previously selected (row,col)
print(col2heading)

i =0
text="mail.com"
for row in range(worksheet.nrows):
    #if str(worksheet.cell(row,7).value) == text:  #Find within Column 8 with Header Domain that matches a text
    if text in str(worksheet.cell(row,7).value) :  #Find within Column 8 with Header Domain that matches a text
        i = i +1  # Increase i by 1 every text matched
        print(worksheet.cell(row,6).value)  # Print the value of the current row and column 7 wich is G
    else:
    	#print(worksheet.cell(row,7).value)
        pass

#print("Found %s" %i)
print("Found %s Email Names that matches the %s domain" %(i,text))