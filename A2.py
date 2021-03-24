# To search for the valid Messstellen_Position in the file
import pandas as pd
file = pd.read_excel(r'C:/Users/Adhwaryu/PycharmProjects/pythonProject/Task2/Task2MINI/Messstellenliste_H2.xlsx')
print(file.head())
print(file.shape)
booleans = []
Messstellen_Position = ["Rohrleitung", "Behälter", "Raum", "Maschine"]
for Messstellen_Position in file.Messstellen_Position:
    if Messstellen_Position == "Rohrleitung":
        booleans.append(True)
    elif Messstellen_Position == "Behälter":
        booleans.append(True)
    elif Messstellen_Position == "Raum":
        booleans.append(True)
    elif Messstellen_Position == "Maschine":
        booleans.append(True)
    else:
        booleans.append(False)
print(booleans[0:5])
len(booleans)
is_position = pd.Series(booleans)
print(is_position.head)
# is_position = file.Messstellen_Position == "Rohrleitung"
# file[file.Messstellen_Position == "Rohrleitung"]
