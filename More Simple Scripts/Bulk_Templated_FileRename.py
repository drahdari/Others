import os
from os import path
import os.path
import sys
from datetime import datetime

print("********************************************")
print("************Bulk File Renaming**************")
print("\n\n\n")

Con="Y"
while Con.upper()=="Y":
    FPath=input("Folder Path: ")
    while not path.exists(str(FPath)):
        FPath=input("Folder Path: ")
    c=1
    for f1 in os.listdir(str(FPath)):
        print(str(c) + " : "+ str(FPath)+"/"+str(f1))
        Date=str(datetime.date(datetime.now()))
        Date=Date.replace("-","")
        Time=str(datetime.time(datetime.now()))
        Time=Time.replace(":","")
        Time,E1,E2=Time.partition('.')
        Name,Dot,Format=f1.partition('.')
        Appending="_ENT_MCT_MCO_002_"+Date+"_"+Time+"."+str(Format)
        os.rename(str(FPath)+"/"+str(f1),str(FPath)+"/"+str(Name)+Appending)
        c+=1
    Con=input("Do You Wish To Continue? (Y/N): ")



