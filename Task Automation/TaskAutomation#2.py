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
from collections import defaultdict 
from operator import itemgetter  
from collections import OrderedDict
import csv
import xlrd
import operator
import re
import collections



COFile = input("Country File Path: ")
Incom_File=input("Incoming File Path:" )
File_Output=input("Output File Path: ")


Country_Dict = defaultdict(list)

workbook = xlrd.open_workbook(COFile)
sheetN=0
row = 2
value=float()
Country_Prefix={}
for sheet in workbook.sheets():
    if sheetN==1:
        while row < 372:
            prefix=str(sheet.cell_value(row,1))
            prefix=prefix[:-2]
            Country_Prefix[str(prefix)]=sheet.cell_value(row,2)
            Country=str(sheet.cell_value(row,2))
            if not Country in Country_Dict:
                Country_Dict[str(sheet.cell_value(row,2))].append(value) 
                Country_Dict[str(sheet.cell_value(row,2))].append(value)
                Country_Dict[str(sheet.cell_value(row,2))].append(value) 
            row +=1
    else:
        sheetN+=1



Numbers= list()
Duration=list()



workbook = xlrd.open_workbook(Incom_File)
First_row=list()
num_rows=-1
for sheet in workbook.sheets():
    num_rows=sheet.nrows-1
    First_row = sheet.row(0)

cl=0
Co_In=-1
Co_F=False
ASR_In=-1
ASR_F=False
NER_In=-1
NER_F=False
Attemp_C=-1
Attempt_F=False

for column in First_row:
    column=str(column)[6::]
    column=str(column)[:-1]
    if str(column) == 'Country':
        Co_In=cl
        Co_F=True
    elif str(column) == 'ASR':
        ASR_In=cl
        ASR_F=True
    elif str(column) == 'NER':
        NER_In=cl
        NER_F=True
    elif str(column) == 'Attempts':
        Attemp_C=cl
        Attempt_F=True
    if NER_F==True and ASR_F==True and Co_F==True and Attempt_F==True:
        break
    cl+=1


row = 1
ASR_Totall=float()
Sum_Attempts=float()
NER_Totall=float()
sheet=workbook.sheets()[0]
Found=False
while(row<int(num_rows)):
        Read_Prefix=str(sheet.cell_value(row,0))
        while not Read_Prefix in Country_Prefix:
            Read_Prefix=Read_Prefix[:-1]
        Country=Country_Prefix[str(Read_Prefix)]
        Read_Name= str(sheet.cell_value(row,Co_In))
        ASR_Value= sheet.cell_value(row,int(ASR_In))
        NER_Value= sheet.cell_value(row,int(NER_In))
        ASR_Value=ASR_Value[:-1]
        NER_Value=NER_Value[:-1]
        print(NER_Value)
        Country_Dict[str(Country)][2] = float(Country_Dict[str(Country)][2]) + float(sheet.cell_value(row,int(Attemp_C)))
        Country_Dict[str(Country)][0] = float(Country_Dict[str(Country)][0]) + float(sheet.cell_value(row,int(Attemp_C))) * float(ASR_Value)
        Country_Dict[str(Country)][1] = float(Country_Dict[str(Country)][1]) + float(sheet.cell_value(row,int(Attemp_C))) * float(NER_Value)
        row+=1


WB = openpyxl.Workbook()
sheet=WB.active
sheet.title = "Average Per Country"
fnt = Font(size=11 , bold= True)

sheet['A1']='Country'
sheet['B1']='ASR Avg'
sheet['C1']='NER Avg'
sheet['D1']='Attempts Sum'


for Country in Country_Dict:
    if not Country_Dict[str(Country)][2] == 0:
        Country_Dict[str(Country)][0] = Country_Dict[str(Country)][0]/Country_Dict[str(Country)][2]
        Country_Dict[str(Country)][1] = Country_Dict[str(Country)][1]/Country_Dict[str(Country)][2]
    str_row = [str(Country), str(Country_Dict[str(Country)][0]), str(Country_Dict[str(Country)][1]), str(Country_Dict[str(Country)][2])]
    sheet.append(str_row)

WB.save(str(File_Output))



