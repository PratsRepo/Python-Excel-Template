# code run sample file
# import os
import pandas as pd
import xlwings as xw
newdir_path = 'C:/Users/Adhwaryu/PycharmProjects/pythonProject/Task2/Task2MINI/Result/'
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
df = pd.DataFrame(Messstellen_data, columns=['Funktion', 'AD65', 'W70', 'B14', 'AC21'])
count_row = df.shape[0]  # Gives number of rows
print("Total rows are: " + str(count_row))
count_col = df.shape[1]  # Gives number of columns
print("Total columns are: " + str(count_col))
# iterate through each row and look for the valid record under Messstellen_Position column
Messstellen_Position = ["Rohrleitung", "Behälter", "Raum", "Maschine"]
booleans = []
# for row in range(1, count_row+1):
# for col in range(1, count_col+1):
for i in df.itertuples(index=True):
    print(i)
    Funktion = df._get_value(0, 'Funktion')
    AD65 = df._get_value(0, 'AD65')
    W70 = df._get_value(0, 'W70')
    B14 = df._get_value(0, 'B14')
    AC21 = df._get_value(0, 'AC21')
    wb = xw.Book(file2)  # open the template file
    sh = wb.sheets[1]  # open the template1 sheet
    print("template1 opened")
    sh.range('X65').value = Funktion[0]
    sh.range('Z65').value = Funktion[1: ]
    sh.range('AD65').value = AD65
    sh.range('W70').value = W70
    sh.range('AC21').value = AC21
    wb.save('file://' + newdir_path + str(AD65) + str("_") + str(Template1) + '.xlsx')
    print("Template printed.")
    booleans.append(True)
    print(booleans[0:15])
    break
