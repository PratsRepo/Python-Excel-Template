# code run sample file -- Working well for iteration of rows and printing template but this is without condition
# import os
import pandas as pd
import xlwings as xw
# from win32com.client import DispatchEx
newdir_path = 'C:/Users/Adhwaryu/PycharmProjects/pythonProject/Task2/Task2MINI/Result/A8/'
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
B14 = ['Rohrleitung', 'Behälter', 'Raum', 'Maschine']
booleans = []
# to iterate all the rows in Messstellenliste file
for i in df.itertuples(index=True):
    # only to read the rows where column AD65(2nd column) does not have blank value or none
    while i[AD65] != 'None':
      # to retrieve the row values and print them in respective template
      # when B14(Messstellen_Position) value is Rohrleitung
      if B14 == 'Rohrleitung':
         print(i)
         # retrieving the values
         Funktion = i.Funktion
         AD65 = i.AD65
         W70 = i.W70
         B14 = i.B14
         AC21 = i.AC21
         wb = xw.Book(file2)  # open the template file
         sh = wb.sheets[0]  # open the template sheet1
         print("template1 opened")
         # mapping the values
         sh.range('X65').value = Funktion[0]
         sh.range('Z65').value = Funktion[1: ]
         sh.range('AD65').value = AD65
         sh.range('W70').value = W70
         sh.range('AC21').value = AC21
         wb.save('file://' + newdir_path + str(AD65) + str("_") + str(Template1) + '.xlsx')
         print("Template printed.")
         wb.close()
         booleans.append(True)
         print(booleans[0:15])
      # to retrieve the row values and print them in respective template
      # when B14(Messstellen_Position) is Behälter
      elif B14 == 'Behälter':
         print(i)
         Funktion = i.Funktion
         AD65 = i.AD65
         W70 = i.W70
         B14 = i.B14
         AC21 = i.AC21
         wb = xw.Book(file2)  # open the template file
         sh = wb.sheets[1]  # open the template sheet2
         print("template1 opened")
         sh.range('X65').value = Funktion[0]
         sh.range('Z65').value = Funktion[1:]
         sh.range('AD65').value = AD65
         sh.range('W70').value = W70
         sh.range('AC21').value = AC21
         wb.save('file://' + newdir_path + str(AD65) + str("_") + str(Template2) + '.xlsx')
         print("Template printed.")
         wb.close()
         booleans.append(True)
      # to retrieve the row values and print them in respective template
      # when B14(Messstellen_Position) is Raum
      elif B14 == 'Raum':
         # repeating the same as above
         # retrieving the values, mapping and saving the template sheet
         booleans.append(True)
      # to retrieve the row values and print them in respective template
      # when B14(Messstellen_Position) is Maschine
      elif B14 == 'Maschine':
          # repeating the same as above
          # retrieving the values, mapping and saving the template sheet
         booleans.append(True)
      else:
    # I want to skip the rows when B14(Messstellen_Position) value is blank






