# -*- coding: utf-8 -*-
import xlsxwriter
import sys, ast

#skrypt pobiera liste tupli z danymi o obecnosci, a potem robi szablon .xlsx na nastepny miesiac
#przykladowe wywolanie skryptu:
#python2 rob_rachunki.py "[['marzec', 'kowalski', ['01.01.2020'], 15, 1, 1, 14, 12]]" "10.19_niepolomice.xlsx"

def dodaj_rachunek (row, col , tup1, worksheet, format1, format2, format3, format4, format5, format6, format7, format8, format9,numer_konta):
    naglowek = "Rachunek - " + tup1[0].decode(encoding='UTF-8')
    nazwisko = tup1[1].decode(encoding='UTF-8')
    worksheet.merge_range(row, col, row, col + 1, naglowek, format1)
    worksheet.merge_range(row + 1, col, row  + 1, col + 1, nazwisko, format6)
    #worksheet.write(row, col, tup1[0].decode(encoding='UTF-8'))
    zajecia = 'Zajęcia:'
    worksheet.write(row + 2, col, zajecia.decode(encoding="UTF-8"), format2)
    worksheet.write(row + 2, col + 1, 'Kwota:', format2)
    row = row + 3
    sum = 0
    for item in tup1[2]:
        worksheet.write(row, col, item, format2)
        if item != '-':
            worksheet.write(row, col + 1, tup1[3], format8)
            sum += tup1[3]
        else:
            worksheet.write(row, col + 1, 0, format8)
        row += 1
    worksheet.write(row, col, 'Suma:', format3)
    znizka = 'Znizka:'
    worksheet.write(row + 1, col, znizka.decode(encoding='UTF-8'), format3)
    worksheet.write(row + 2, col, 'Zwrot:', format3)
    zaleglosci = 'Zaległości:'
    worksheet.write(row + 3, col, zaleglosci.decode(encoding='UTF-8'), format3)
    worksheet.write(row, col + 1, sum, format8)
    worksheet.write(row + 1, col + 1, 10*tup1[4], format7)
    worksheet.write(row + 2, col + 1, tup1[3]*tup1[5]+tup1[6], format8)
    worksheet.write(row + 3, col + 1, tup1[7], format8)
    kwota = 'KWOTA DO ZAPŁATY:'
    worksheet.write(row + 4, col, kwota.decode(encoding='UTF-8'), format4)
    worksheet.write(row + 4, col +1, (1 - 0.1*tup1[4]) * (sum-tup1[5]*tup1[3]) - tup1[6] + tup1[7], format9)
    worksheet.merge_range(row + 5, col, row + 5, col + 1, 'Dane do przelewu:', format2)
    worksheet.write(row + 6, col, 'Odbiorca:', format5)
    wlascicielka='Liliana Wrońska'
    worksheet.write(row + 6, col + 1, wlascicielka.decode(encoding='UTF-8'), format2)
    worksheet.merge_range(row + 7, col, row + 7, col + 1, 'Numer konta: '+numer_konta, format2)

def ustaw_wiersze(r, worksheet, l_dni):
    # rozmiar wierszy
    worksheet.set_row(r + 0, 20.7)
    worksheet.set_row(r + 1, 20.7)
    worksheet.set_row(r + 2, 15.4)
    for i in range(0, l_dni): # 0, 1, ..... l_dni - 1
        worksheet.set_row(r + 3 + i, 15.6 * 4/l_dni) #skalowanie: optymalny rozmiar to 15.6 dla 4 dni
    worksheet.set_row(r + i + 4, 14.3)
    worksheet.set_row(r + i + 5, 14.3)
    worksheet.set_row(r + i + 6, 14.3)
    worksheet.set_row(r + i + 7, 14.3)
    worksheet.set_row(r + i + 8, 19.8)
    worksheet.set_row(r + i + 9, 15.4)
    worksheet.set_row(r + i + 10, 15.4)
    worksheet.set_row(r + i + 11, 15.4)
    worksheet.set_row(r + i + 12, 1)

def zmien_nazwe(nazwa):
    miesiac = nazwa[0] + nazwa[1]
    rok = nazwa[3] + nazwa[4]
    if int(miesiac) <= 11:
        miesiac = str(int(miesiac) + 1)
        if int(miesiac) <= 9:
            miesiac = "0" + miesiac
    else:
        rok = str(int(rok) + 1)
        miesiac = "01"
    nowa = miesiac + "." + rok + nazwa[5:]
    return nowa

def przestaw_miesiac(miesiac):
    miesiace = ['styczeń','luty','marzec','kwiecień','maj','czerwiec','lipiec','sierpień','wrzesień','październik','listopad','grudzień']
    try: 
        i = miesiace.index(miesiac)
        i = (i + 1)%len(miesiace)
        return miesiace[i]
    except Exception as e:
        raise Exception("Nie znaleziono "+miesiac+" na liście miesięcy.")

#PROGRAM

#parametry
#prawdziwe odczytanie hasla przebiega inaczej
numer_konta = "123456789"

#dane
# [[miesiac,nazwisko,[daty],cena_zajec, z,p,zwrot_ekstra,zaleglosc,email],[]...]
inputList = ast.literal_eval( sys.argv[1] )

nazwapliku = sys.argv[2]
nazwapliku = zmien_nazwe(nazwapliku)

#tworzenie arkusza
workbook = xlsxwriter.Workbook(nazwapliku)
worksheet = workbook.add_worksheet('rachunki')

#formaty
napis_waluta = '# ##0.00 [$zł-415]'
format1 = workbook.add_format({'font_size': 13, 'align': 'center','bold': True, 'border': 1, 'valign': 'vcenter'})
format2 = workbook.add_format({'font_size': 11, 'align': 'center', 'border': 1, 'valign': 'vcenter'})
format6 = workbook.add_format({'font_size': 12, 'align': 'center', 'border': 1, 'valign': 'vcenter'})
format3 = workbook.add_format({'font_size': 11, 'align': 'right', 'border': 1, 'valign': 'vcenter'})
format4 = workbook.add_format({'font_size': 12, 'align': 'center','bold': True, 'border': 1, 'valign': 'vcenter'})
format5 = workbook.add_format({'font_size': 11, 'align': 'left', 'border': 1, 'valign': 'vcenter'})
format7 = workbook.add_format({'font_size': 12, 'align': 'center', 'border': 1, 'num_format': '0"%"', 'valign': 'vcenter'})
format8 = workbook.add_format({'font_size': 11, 'align': 'center', 'border': 1, 'num_format': napis_waluta.decode(encoding='UTF-8'), 'valign': 'vcenter'})
format9 = workbook.add_format({'font_size': 12, 'align': 'center','bold': True, 'border': 1, 'num_format':'#,##0.00 "zl"', 'valign': 'vcenter'})

#rozmiar kolumn
worksheet.set_column('A:D', 21)

#zapis tabeli danych
r = 0
c = 2
i = 0
l_dni = len(inputList[0][2])
for item in inputList:
    if i%2 == 0:
        ustaw_wiersze(r, worksheet, l_dni)
        c = c - 2 #przesuwa na pozycje lewej kolumny
    else:
        c = c + 2
    item[0] = przestaw_miesiac(item[0])#zmienia na nast miesiac
    dodaj_rachunek(r, c, item, worksheet, format1, format2, format3, format4, format5, format6, format7, format8, format9,numer_konta)
    if i%2 == 1:
        r = r + 12 + l_dni
    i = i + 1

workbook.close()
