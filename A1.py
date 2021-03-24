# code run sample file
# import os
import pandas as pd
import xlwings as xw
file1 = "Messstellenliste_H2.xlsx"
file2 = "Template.xlsx"
Template1 = "Messstelle_Rohrleitung"
Template2 = "Messstelle_Behälter"
Template3 = "Messstelle_Maschine"
Template4 = "Messstelle_Raum"
fileext = ".xlsx"
# read the source file
Messstellen_data = pd.read_excel(file1, sheet_name=0, header=0,
                                 index_col=False, keep_default_na=True)
# creating pandas dataframe from the source file
df = pd.DataFrame(Messstellen_data, columns=['Function', 'Zahlnummer', 'Benennung', 'Messstellen_Position',
                                             'EzANummer'])
# print(df)  # print the dataframe
# print(df.head())  #print top 5 rows of the table
# print(df.shape)
# print(df.count)
count_row = df.shape[0]  # Gives number of rows
print("Total rows are: " + str(count_row))
count_col = df.shape[1]  # Gives number of columns
print("Total columns are: " + str(count_col))
# iterate through each row and look for the valid record under Messstellen_Position column
Messstellen_Position = ["Rohrleitung", "Behälter", "Raum", "Maschine"]
booleans = []
# for row in range(df.shape[0],):
    # for col in range(df.shape[1]):
for row in range(1, count_row+1):
    for col in range(1, count_col+1):
        for Messstellen_Position in df.Messstellen_Position:
            if Messstellen_Position == "Rohrleitung":
                # we use iteritems() function this function iterates over each column as key,
                # value pair with label as key and column value as a Series object
                for i, j in df.iterrows():
                    # read the row and get the values from all the columns
                    print(i, j)
                    i = i + 1
                    print()
                # subdivide the value of funktion into funktion and stellenfunktion
                wb = xw.Book(file2)   # open the template file
                sh = wb.sheets[0]       # open the template1 sheet
                print("template1 opened")
                booleans.append(True)
                break
            elif Messstellen_Position == "Behälter":
                for i in df.itertuples():
                    print(i)
                wb = xw.Book(file2)  # open the template file
                sh = wb.sheets[1]  # open the template1 sheet
                print("template2 opened")
                booleans.append(True)
            elif Messstellen_Position == "Raum":
                for key, value in df.iteritems():
                    # read the row and get the values from all the columns
                    print(key, value)
                    print()
                wb = xw.Book(file2)  # open the template file
                sh = wb.sheets[2]  # open the template1 sheet
                print("template3 opened")
                booleans.append(True)
            elif Messstellen_Position == "Maschine":
                wb = xw.Book(file2)  # open the template file
                sh = wb.sheets[3]  # open the template1 sheet
                print("template4 opened")
                booleans.append(True)
            else:
                booleans.append(False)
    print(booleans[0:15])
    break
