import xlwings as xw
import os
# from win32com.client import DispatchEx
cwd = os.path.abspath(r'C:/Users/Adhwaryu/PycharmProjects/pythonProject/Task2')
newdir_path = 'C:/Users/Adhwaryu/PycharmProjects/pythonProject/Task2/FinalTemplateFiles/'
file1 = "Messstellenliste_H2 zug_FEST.xlsx"
file2 = "BLANCO_encrypted.xlsx"
Template1 = "Messstelle_Rohrleitung"
Template2 = "Messstelle_Behälter"
Template3 = "Messstelle_Maschine"
Template4 = "Messstelle_Raum"
fileext = ".xlsx"

# if Messstellen-Position = = "Rohrleitung"
# then
# open the main template file
# wb = xw.Book(file)
# put logic to search for the Messstelle type and then call for the corresponding sheet ti write
# sh = wb.sheets[0]
# elif Messstellen-Position = = "Behälter"
# open the main template file
# wb = xw.Book(file)
# put logic to search for the Messstelle type and then call for the corresponding sheet ti write
# sh = wb.sheets[1]
# elif Messstellen-Position = = "Maschine"
# open the main template file
# wb = xw.Book(file)
# put logic to search for the Messstelle type and then call for the corresponding sheet ti write
# sh = wb.sheets[2]
# else: Messstellen-Position = = "Raum"
# open the main template file
# wb = xw.Book(file)
# put logic to search for the Messstelle type and then call for the corresponding sheet ti write
# sh = wb.sheets[3]

# open the Messstellenliste file
wb = xw.Book(file1)
sh = wb.sheets[0]

# Retrieving values from sheet1 and mapping into the fields
Funktion = sh.range('A10').value
print("Value is:", Funktion)

neueTAGNr = sh.range('B10').value
print("Value is:", neueTAGNr)

Benennung = sh.range('D10').value
print("Value is:", Benennung)

Bezeichnung = sh.range('AB10').value
print("Value is:", Bezeichnung)

Zusammensetzung = sh.range('AC10').value
print("Value is:", Zusammensetzung)

Korrosive_Bestandteile = sh.range('AE10').value
print("Value is:", Korrosive_Bestandteile)

Schwebstoffe = sh.range('AC10').value
print("Value is:", Schwebstoffe)

# open the main template file
wb = xw.Book(file2)
# put logic to search for the Messstelle type and then call for the corresponding sheet to write
sh = wb.sheets[0]
sh.range('J2').value = Bezeichnung
sh.range('J3').value = Zusammensetzung
sh.range('J4').value = Korrosive_Bestandteile
sh.range('J5').value = Schwebstoffe
sh.range('AD65').value = neueTAGNr
sh.range('W70').value = Benennung
Zahlnummer = sh.range('AD65').value
# save the file to the new directory with the new tag number and the corresponding template file name

wb.save('file://' + newdir_path + str(Template1) + str("_") + str(Zahlnummer) + '.xlsx')
print("Template printed.")
