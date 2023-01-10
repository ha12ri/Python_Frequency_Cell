# import xlrd

# # Taking the user input for the word 
# word = input("Enter any word : ") 

# # Variable for storing the frequency ( count of occurences )
# frequency  = 0

# # loading the workbook
# workbook = xlrd.open_workbook('FrequencyExcel\SampleData.xlsx')

# # Storing the sheet 1 of the workbook 
# # because sheet 1 contains our orignal data that we will use in our program
# sheet_one = workbook.sheet_by_index(0)

# # You may use below code to check the value in cell[5][4]
# # print("Value of 5-4 cell: ",sheet_one.cell_value(5, 4))

# # Get  no of rows and columns
# # print("Number of Rows: ", sheet_one.nrows)
# # print("Number of Columns: ",sheet_one.ncols)

# # Traversing the rows using outer for loop
# for row in range (sheet_one.nrows):
#     # Traversing the columns using inner for loop
#     for col in range (sheet_one.ncols):
#         # Checking if any cell value equals to our input word
#         if(sheet_one.cell_value(row,col) == word):
#             # If true incrementing frequency count 
#             frequency = frequency + 1
            
            
# # Printing the frequency of the word 
# print("The frequency of the given word : " + word + " is : ", frequency)


