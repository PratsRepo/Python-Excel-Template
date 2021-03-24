# code run sample file
import os
import pandas as pd
import xlwings as xw
cwd = os.path.abspath(r'C:/Users/Adhwaryu/PycharmProjects/pythonProject/Task2')
newdir_path = 'C:/Users/Adhwaryu/PycharmProjects/pythonProject/Task2/FinalTemplateFiles/'
file1 = "Messstellenliste_H2.xlsx"
file2 = "BLANCO_Template.xlsx"
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
# iterate through each row
start_row = df.head(1)
start_column = df.iloc[:, 0]
# print("First row is: " + str(start_row))
total_row = df.shape[0]  # Gives number of rows
print("Total rows are: " + str(total_row))
total_col = df.shape[1]  # Gives number of columns
print("Total columns are: " + str(total_col))


def loop_through_excel(start_row, total_row):
    row_cursor_found = 0
    for row_cursor in range(start_row, start_column):
        function = df.cell(row_cursor_found, 1).value
        print(function)
    for i in df.itertuples():
        print(i)


# setup the condition for Messstellen_Position
start_row = 1
Messstellen_Position = ["Rohrleitung", "Behälter", "Raum", "Maschine"]

while Messstellen_Position == "Rohrleitung":
    Messstellen_Position = loop_through_excel(start_row, total_row)

    if Messstellen_Position == "Rohrleitung":
        for i in df.itertuples():
            df.get(i)
            print(i)
            wb = xw.Book(file2)  # open the template file
            sh = wb.sheets[0]    # open the template1 sheet
            df.set_index()       # copying the values in the template sheet
            sh.range('X65').value = Funktion
            sh.range('Z65').value = Stellenfunktion
            sh.range('AD65').value = Zahlnummer
            sh.range('W70').value = Benennung
            sh.range('AC21').value = EzANummer
            # save the file to the new directory with the new tag number
            # and the corresponding template file name
            wb.save('file://' + newdir_path + str(Template_Name) + str("_") + str(Zahlnummer) + '.xlsx')
            print("Template printed.")
            break
