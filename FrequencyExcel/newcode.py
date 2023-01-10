# Importing the modules
import pandas as pd
from operator import * 

# Taking input for the word
word = str(input("Enter any word : "))

# Loading the Excel sheet into pandas DataFrame
Xcel_file = pd.read_excel('FrequencyExcel\sampledatafoodsales.xlsx')

# print(Xcel_file) (Optional)

# a = Xcel_file['Category'].value_counts()

# print(a)

# frequecy count 
freq = 0

# freq = countOf(Xcel_file['Category'], "Bars")
# print(freq)

# Traversing each column
for i in Xcel_file:
    # print(i, '\n')
    
    # freuency in ith column
    c = countOf(Xcel_file[i], word)
    
    # incrementing the final freuency with ith column freuency
    freq +=c

# printing the final frequency in Excel table 
print(freq)


    
