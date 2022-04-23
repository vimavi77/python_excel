import pandas as pd
file="/Users/macbook/Desktop/excel_python/reporte.xlsx"
file_output="/Users/macbook/Desktop/excel_python/test.xlsx"
#codigo creado 21-mar-2022
#xls = pd.ExcelFile(file) # Read the Excel file
# Now you can list all sheets in the file
#print(xls.sheet_names)
#for sheet_name in xls.sheet_names:
#	print(sheet_name)
df = pd.read_excel(file,sheet_name=None) # Read the whole file with sheets names into a Data Frame
#print(df.keys())  # Print the whole Sheets names in the Data Frame
i=0
for sheet_name in df.keys():
	i=i+1
	print(sheet_name)
	df[sheet_name].sort_values(by=['EQUIPO'], inplace=True)
	if i == 1 : # Crear archivo con primera Sheet
		df[sheet_name].to_excel (file_output, index = False, header=True,sheet_name=sheet_name)
	else : # Hacer append al archivo creado con el Sheet nuevo
		with pd.ExcelWriter(file_output,mode='a') as writer: 
			df[sheet_name].to_excel(writer, sheet_name=sheet_name,index = False)
#df['TIJUANA'].sort_values(by=['Equipo'], inplace=True) # Sort ascending order inplace has to be 'True' 'False' is for descending
#Note that unless specified, the values will be sorted in an ascending order by default.
#print(df)  # Print the whole Data Frame
#df['TIJUANA'].to_excel ('/Users/macbook/Desktop/excel_python/test.xlsx', index = False, header=True,sheet_name='Your sheet name')