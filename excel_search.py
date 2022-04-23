#codigo creado 1-may-2021
import pandas as pd
import numpy as np

excel_files=['/Users/macbook/Desktop/remove_duplicates.xlsx']

i =0
text="gmail.com"  # The text value to be matched
for file in excel_files:
	df = pd.read_excel(file) # Create a Data Frame by reading the excel file
	#email=df['Email Name'].where(df['Domain']=='gmail.com').dropna() #Search in header Email Name wich has the value gmail.com in Colum Domain
	email=df['Email Name'].where(df['Domain']==text).dropna() #Search in header Email Name wich has the value gmail.com in Colum Domain
	print(email)  # Print the email that matched the text
	print(df)  # Print the whole Data Frame
print("Found Email Names that matches the %s domain" %(text))