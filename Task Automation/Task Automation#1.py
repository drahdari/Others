import os
import sys
from datetime import date
import smtplib, ssl
import email.utils
import openpyxl
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email import encoders
from openpyxl import workbook
from openpyxl.styles import Color, Font
from openpyxl.styles import Alignment
from openpyxl.styles.colors import GREEN
import csv
import xlrd
import operator
import re



FFile= input("First File Path: ")
COFile = input("Country Code File Path: ")
File_Output=input("Output File: ")


Numbers= list()
Duration=list()

Country_Dict = {}

workbook = xlrd.open_workbook(COFile)
sheetN=0
row = 2
CountryName=list()
CountryPrefix=list()
for sheet in workbook.sheets():
    if sheetN==1:
        while row < 372:
            Country_Dict[str(sheet.cell_value(row,2))]=0
            CountryName.append(str(sheet.cell_value(row,2)))
            CountryPrefix.append(sheet.cell_value(row,1))
            row +=1
    else:
        sheetN+=1
Numbers= list()
Duration=list()

with open(FFile, newline='') as csvfile:
     reader = csv.reader(csvfile, delimiter=';')
     c=0
     TPI=-1
     DuI=-1
     for row in reader:
        if c!=0:
            if len(str(row[DuI])) != 0:
                if len(str(row[TPI])) !=0:
                    Numbers.append(row[TPI])
                else:
                    Numbers.append("N/A")
                Duration.append(row[DuI])
        else:
            l=0
            for item in row:
                if item == "TP ANI":
                    TPI=l
                elif item =="Duration":
                    DuI=l
                l+=1
            c+=1


c1=0
CountryPrefix2=list()
for prefix in CountryPrefix:
    prefix=str(prefix).strip('.0')
    prefix= "+" + str(prefix)
    CountryPrefix2.append(str(prefix))

CountryPrefix=CountryPrefix2



EndResult=list()
EndName=list()

for number in Numbers:
    c2=0
    Find=False
    if str(number) == "N/A" or str(number) == "anonymous" or len(str(number)) < 3 or str(number)=="0000" or str(number)=="TP ANI" or not str(number)[1].isdigit():
        EndName.append("N/A")
    else:

        match=list()
        match_index=list()
        for Prefix in CountryPrefix:
            if (number.startswith("+")):
                number = str(number)[1 : : ]
            while str(number).startswith("0"):
                number = str(number)[1 : : ]
            if not str(number).startswith("+"):
                number="+" + number
            if str(number).startswith(str(Prefix)):
                match.append(Prefix)
                match_index.append(c2)
                Find=True
            c2+=1
        if Find==False:
            print(number) 
            EndName.append("Unknown")
        else:
            longest_match=-1
            longest_lenght=-1
            lonest_index=-1
            c=0
            for item in match:
                if len(item)>longest_lenght:
                    longest_lenght=len(item)
                    longest_index=match_index[c]
                    longest_match=item
                c+=1
            EndName.append(CountryName[int(longest_index)])

        
WOrksheet_NAMES=["Detailed Usage","Duration By Country"]
WB = openpyxl.Workbook()
sheet=WB.active
sheet.title = "Detailed Usage"
fnt = Font(size=11 , bold= True)

sheet['A1']='TP ANI'
sheet['B1']='Country Code'
sheet['C1']='Duration'
c=0
for number in Numbers:
    str_row = [str(number), str(EndName[c]), str(Duration[c])]
    sheet.append(str_row)
    c+=1


c=0
for Country in EndName:
    if Country!="N/A" and Country!="Unknown":
        print(Country)
        Country_Dict[str(Country)] = int(Country_Dict[str(Country)]) + int(Duration[c]) 
    c+=1 

WB.create_sheet("Duration By COuntry")
WB.active=1
sheet=WB.active

sheet['A1']='Country Name'
sheet['B1']='Duration'

Country_Dict=sorted(Country_Dict.items(), key = operator.itemgetter(1), reverse=True)


for Country, time in Country_Dict:
    time = float(time)/60
    str_row= [str(Country), str(time)]
    sheet.append(str_row)

WB.save(str(File_Output))



