# -*- coding: utf-8 -*-

#skrypt wysyła maile z rachunkami na podstawie podanych parametrow
#przykladowe wywolanie skryptu:
#python2 wyslij_maile.py "[['marzec','kowalski',['01.01.2020'],15,1,1,0,0,'mail@mail.com']]" "10.19_niepolomice.xlsx"

import xlsxwriter
import sys, ast
import smtplib

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

def wyslij_rachunek (tup, nadawca, haslo, miejscowoscimiesiac, numer_konta):
    #[[miesiac,nazwisko,[daty],cena_zajec, z,p,zwrot_ekstra,zaleglosc,email],[]...]
    gmail_user = nadawca
    gmail_password = haslo
    email_from = nadawca
    email_to = tup[8]
    email_subject = "Szkółka tańca TAKT - rachunek za " + przestaw_miesiac(tup[0])
    
    email_body = tup[1] + "\n" + "zajęcia:\tkwota:\n"
    suma = 0
    for data in tup[2]:
        if data == "-":
            email_body = email_body + data + "\t" + "0zł" + "\n"
        else:
            email_body = email_body + data + "\t" + str(tup[3]) + "zł\n"
            suma += tup[3]
    email_body = email_body + "\n"
    email_body = email_body + "suma:\t" + str(suma) + "zł\n"
    email_body = email_body + "zniżka:\t" + str(10*tup[4]) + "%\n"
    email_body = email_body + "zwrot:\t" + str(tup[3]*tup[5]+tup[6]) + "zł\n"
    email_body = email_body + "zaległości:\t" + str(tup[7]) + "zł\n"
    email_body = email_body + "kwota do zapłaty:\t" + str((1 - 0.1*tup[4]) * (suma-tup[5]*tup[3]) - tup[6] + tup[7]) + "zł\n"
    email_body = email_body + "\n"

    email_body = email_body + "tytuł przelewu:\t" + miejscowoscimiesiac + " " + tup[1] + "\n"
    email_body = email_body + "odbiorca:\t" +  "TAKT Liliana Wrońska\n"
    email_body = email_body + "numer konta:\t" +  numer_konta + "\n"
    #termin płatności
    
    email_text = """\
Subject: %s
   
%s
""" % (email_subject,email_body)
    try:  
        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        server.login(gmail_user, gmail_password)
        server.sendmail(email_from, email_to, email_text)
        server.close()
    except Exception as E:  
        print 'Something went wrong while sending email...'
        print E

#parametry
#prawdziwe odczytanie hasla przebiega inaczej
nadawca = "mail@mail.com"
haslo = "123456789"
numer_konta = "123456789"

#dane
# [[miesiac,nazwisko,[daty],cena_zajec, z,p,zwrot_ekstra,zaleglosc,email],[]...]
inputList = ast.literal_eval( sys.argv[1] )

# '10.19_niepolomice.xlsx'
nazwapliku = sys.argv[2]
nazwapliku = zmien_nazwe(nazwapliku)
miejscowoscimiesiac = nazwapliku.partition(".xlsx")[0]

for item in inputList:
    wyslij_rachunek(item,nadawca,haslo,miejscowoscimiesiac,numer_konta)

