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

# open the Messstellenliste file
wb = xw.Book(file1)
sh = wb.sheets[0]

Messstellenliste = []

# Retrieving values and mapping into the fields
# common fields to all template
# Map Funktion data value to Funktion and Stellenfunktion fields
Function = sh.range('A12').value
Funktion = Function[0]
print("Funktion Value is:", Funktion)
Stellenfunktion = Function[1: ]
print("Stellenfunktion Value is:", Stellenfunktion)

Zahlnummer = sh.range('B12').value
print("Value is:", Zahlnummer)

Benennung = sh.range('D12').value
print("Value is:", Benennung)
Messstellen_Position = sh.range('D12').value
print("Value is:", Messstellen_Position)

Fliebild_Nr = sh.range().value
# Geräte Angaben
Messbereich_Einheit = sh.range().value
Messbereich_min = sh.range().value
Messbereich_max = sh.range().value
Messbereichsgrenzen_Einheit = sh.range().value
Messbereichsgrenzen_min = sh.range().value
Messbereichsgrenzen_max = sh.range().value
Umgebungstemperatur_min = sh.range().value
Umgebungstemperatur_max = sh.range().value
Ausgangssignal = sh.range().value
Hilfsenergie = sh.range().value
SIL_Kategorie = sh.range().value
Schutzart = sh.range().value
Explosionsschutz = sh.range().value
EG_Baumuster = sh.range().value
Zeugnisse = sh.range().value
Bescheinigungen = sh.range().value
Gerate_Bemerkung = sh.range().value
Montagehinweis1 = sh.range().value
# Medium
Bezeichnung = sh.range().value
Zusammensetzung = sh.range().value
Korrosive_B = sh.range().value
Schwebstoffe = sh.range().value
Arbeitstemperatur_Einheit = sh.range().value
Arbeitstemperatur_min = sh.range().value
Arbeitstemperatur_norm = sh.range().value
Arbeitstemperatur_max = sh.range().value
Arbeitsdruck_Einheit = sh.range().value
Arbeitsdruck_min = sh.range().value
Arbeitsdruck_norm = sh.range().value
Arbeitsdruck_max = sh.range().value
Durchfluss_Einheit = sh.range().value
Durchfluss_min = sh.range().value
Durchfluss_norm = sh.range().value
Durchfluss_max = sh.range().value
pHWert_Einheit = sh.range().value
pHWert_min = sh.range().value
pHWert_norm = sh.range().value
pHWert_max = sh.range().value
DyVisko_Einheit = sh.range().value
DyVisko_min = sh.range().value
DyVisko_norm = sh.range().value
DyVisko_max = sh.range().value
Dichte_Einheit = sh.range().value
Dichte_min = sh.range().value
Dichte_norm = sh.range().value
Dichte_max = sh.range().value
# Anschluss
Stutzenlange = sh.range().value
Typ = sh.range().value
Werkstoff = sh.range().value
Nennweite = sh.range().value
Nenndruck = sh.range().value
# Umgebung
Umge_Bezeichnung = sh.range().value
ExBedingungen = sh.range().value
Sonstige_Auflagen = sh.range().value
Umge_temperatur = sh.range().value
Umgebungsdruck = sh.range().value
# PLT-Stelle
Qualitats = sh.range().value
Sicherheits = sh.range().value
GMP = sh.range().value
Klassifizierung = sh.range().value
EzAPLT = sh.range().value
EzANummer = sh.range().value
Prufintervall = sh.range().value
Verarbeitungsort = sh.range().value
NadEs = sh.range().value
# Rohrleitung
Rohrleitungsken = sh.range().value
Rohr_Nennweite = sh.range().value
Rohr_Nenndruck = sh.range().value
Rohr_ausen = sh.range().value
Rohr_innen = sh.range().value
Rohr_Wanddicke = sh.range().value
Rohr_Werkstoff = sh.range().value
Isolationsstarke = sh.range().value
Rohr_HeizKuhlung = sh.range().value
# Behälter
Apparatekurzzeichen = sh.range().value
Beha_Kommentar = sh.range().value
Beha_Hohe = sh.range().value
Beha_innen = sh.range().value
Beha_Wanddicke = sh.range().value
Beha_Werkstoff = sh.range().value
Volumen = sh.range().value
Beha_HeizKuhlung = sh.range().value
# Raum
Raum_Name = sh.range().value
Raum_Kommentar = sh.range().value
Raum_Aufgabe = sh.range().value
Raum_Etage = sh.range().value
Raum_Bemerkung1 = sh.range().value
Raum_Bemerkung2 = sh.range().value
Maschine_Name = sh.range().value
# Maschine
Maschine_Kommentar = sh.range().value
Maschine_Aufgabe = sh.range().value
Maschine_Hersteller = sh.range().value
Maschine_Typbezeichnung = sh.range().value
Maschine_Bemerkung1 = sh.range().value
Maschine_Bemerkung2 = sh.range().value

# open the main template file
wb = xw.Book(file2)
# put logic to search for the Messstelle type and then call for the corresponding sheet ti write
sh = wb.sheets[0]
# Mapping the data under Medium section
sh.range('J2').value = Bezeichnung
sh.range('J3').value = Zusammensetzung
sh.range('J4').value = Korrosive_B
sh.range('J5').value = Schwebstoffe
sh.range('J7').value = Arbeitstemperatur_Einheit
sh.range('M7').value = Arbeitstemperatur_min
sh.range('P7').value = Arbeitstemperatur_norm
sh.range('R7').value = Arbeitstemperatur_max
sh.range('J9').value = Arbeitsdruck_Einheit
sh.range('M9').value = Arbeitsdruck_min
sh.range('P9').value = Arbeitsdruck_norm
sh.range('R9').value = Arbeitsdruck_max
sh.range('J10').value = Durchfluss_Einheit
sh.range('M10').value = Durchfluss_min
sh.range('P10').value = Durchfluss_norm
sh.range('R10').value = Durchfluss_max
sh.range('J11').value = pHWert_Einheit
sh.range('M11').value = pHWert_min
sh.range('P11').value = pHWert_norm
sh.range('R11').value = pHWert_max
sh.range('J12').value = DyVisko_Einheit
sh.range('M12').value = DyVisko_min
sh.range('P12').value = DyVisko_norm
sh.range('R12').value = DyVisko_max
sh.range('J13').value = Dichte_Einheit
sh.range('M13').value = Dichte_min
sh.range('P13').value = Dichte_norm
sh.range('R13').value = Dichte_max

# Mapping the data under Anschluß section
sh.range('J21').value = Stutzenlange
sh.range('J22').value = Typ
sh.range('J23').value = Werkstoff
sh.range('J24').value = Nennweite
sh.range('J25').value = Nenndruck

# Mapping the data under Umgebung section
sh.range('J26').value = Umge_Bezeichnung
sh.range('J27').value = ExBedingungen
sh.range('J28').value = Sonstige_Auflagen
sh.range('J29').value = Umge_temperatur
sh.range('J30').value = Umgebungsdruck

# Mapping the data under Geräte Angaben section
sh.range('AC3').value = Messbereich_Einheit
sh.range('AE3').value = Messbereich_min
sh.range('AG3').value = Messbereich_max
sh.range('AC4').value = Messbereichsgrenzen_Einheit
sh.range('AE4').value = Messbereichsgrenzen_min
sh.range('AG4').value = Messbereichsgrenzen_max
sh.range('AE5').value = Umgebungstemperatur_min
sh.range('AG5').value = Umgebungstemperatur_max
sh.range('AC6').value = Ausgangssignal
sh.range('AC7').value = Hilfsenergie
sh.range('AC9').value = SIL_Kategorie
sh.range('AC10').value = Schutzart
sh.range('AC11').value = Explosionsschutz
sh.range('AC12').value = EG_Baumuster
sh.range('AC13').value = Zeugnisse
sh.range('AC14').value = Bescheinigungen
sh.range('AC15').value = Gerate_Bemerkung
sh.range('AC16').value = Montagehinweis1

# Mapping the data under PLT-Stelle section
sh.range('AC18').value = Qualitats
sh.range('AE18').value = Sicherheits
sh.range('AG18').value = GMP
sh.range('AC19').value = Klassifizierung
sh.range('AC20').value = EzAPLT
sh.range('AC21').value = EzANummer
sh.range('AC22').value = Prufintervall
sh.range('AC23').value = Verarbeitungsort
sh.range('AC24').value = NadEs

# Mapping the data under Rohrleitung section
sh.range('J14').value = Rohrleitungsken
sh.range('J15').value = Rohr_Nennweite
sh.range('P15').value = Rohr_Nenndruck
sh.range('M17').value = Rohr_ausen
sh.range('P17').value = Rohr_innen
sh.range('R17').value = Rohr_Wanddicke
sh.range('J18').value = Rohr_Werkstoff
sh.range('J19').value = Isolationsstarke
sh.range('J20').value = Rohr_HeizKuhlung

# Mapping the data under Behälter section
sh.range('J14').value = Apparatekurzzeichen
sh.range('J15').value = Beha_Kommentar
sh.range('M17').value = Beha_Hohe
sh.range('P17').value = Beha_innen
sh.range('R17').value = Beha_Wanddicke
sh.range('J18').value = Beha_Werkstoff
sh.range('J19').value = Volumen
sh.range('J20').value = Beha_HeizKuhlung

# Mapping the data under Raum section
sh.range('J14').value = Raum_Name
sh.range('J15').value = Raum_Kommentar
sh.range('J16').value = Raum_Aufgabe
sh.range('J17').value = Raum_Etage
sh.range('J18').value = Raum_Bemerkung1
sh.range('J19').value = Raum_Bemerkung2

# Mapping the data under Maschine section
sh.range('J14').value = Maschine_Name
sh.range('J15').value = Maschine_Kommentar
sh.range('J16').value = Maschine_Aufgabe
sh.range('J17').value = Maschine_Hersteller
sh.range('J18').value = Maschine_Typbezeichnung
sh.range('J19').value = Maschine_Bemerkung1
sh.range('J20').value = Maschine_Bemerkung2

# Mapping the data under Zug. PLT-Stellen section
# Mapping the data under Signale section
# Mapping the data under Betriebsmittel section
# Mapping the data under Bemerkungen section
# Mapping the data under Grafik section #B63-O67

# Mapping the data under PLT-StelLe section
sh.range('X65').value = Funktion
sh.range('Z65').value = Stellenfunktion
sh.range('AD65').value = Zahlnummer
sh.range('W70').value = Benennung
# Zahlnummer = sh.range('AD65').value
# sh.range('F73').value = Projekt
# sh.range('F74').value = Gebäude
sh.range('AA73').value = Fliebild_Nr
# sh.range('AA74').value = Anlagenbezeichner

# save the file to the new directory with the new tag number and the corresponding template file name

wb.save('file://' + newdir_path + str(Template1) + str("_") + str(Zahlnummer) + '.xlsx')
print("Template printed.")
