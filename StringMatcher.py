#!/bin/python

## CODED in Python 3.7 ##

print ("\n+-----------+-----------+")
print (" String Match Automation \n   by Mark Mon Monteros")
print ("+-----------+-----------+")
print (">> Coded in Python 3.7 <<\n")

import subprocess
import sys
import os

matchFile = subprocess.check_output("cat ~/Desktop/'Supply and Demand Tool.csv' | tr ' ' '_' | awk '{print$1}'", shell=True).split()

###ADD ROLE HERE
roles = "Application Support Engineer\\|"
roles += "Project Control Services Practitioner\\|"
roles += "Application Developer\\|"
roles += "Application Lead\\|"
roles += "Advance Application Engineer\\|"
roles += "Tester\\|"
roles += "Application Designer\\|"
roles += "Program/Project Manager\\|"
roles += "Mobilization Lead"

###USING DEMAND-UPSERT CSV FILE
shellCommand1 = "cat ~/Desktop/demand-upsert05042018.csv | awk -F \"\\\"*,\\\"*\" '{print$21 " " $22 " " $23 " " $24 " " $25 " " $26 " " $27}' "
shellCommand1 += "| grep -o 'R[0-9]*' | grep -v R$"
shellCommand2 = "cat ~/Desktop/demand-upsert05042018.csv | grep weeks "
shellCommand2 += "| grep -o " + "\"" + roles + "\""
shellCommand2 += " | tr ' ' '_'"

mainSourceCol1 = subprocess.check_output(shellCommand1, shell=True).split()     #Assign RRD Column
mainSourceCol2 = subprocess.check_output(shellCommand2, shell=True).split()     #Assign Role Column

###REDIRECTING OUTPUT TO DRAFT.CSV
path = "c:\\%HOMEPATH%\\Desktop\\"

with open("draft.csv", "w") as f:
    print("Filename: ", path, file=f)

orig_stdout = sys.stdout
f = open("draft.csv", "w")
sys.stdout = f

for i in range(1,len(matchFile)):
        for j in range(0,len(mainSourceCol1)):
                if (matchFile[i]==mainSourceCol1[j]):
                        print(matchFile[i],mainSourceCol2[j])

sys.stdout = orig_stdout
f.close()

###FINAL OUTPUT FILE
fileName = input("Enter filename: ")
with open(fileName + ".csv", "w") as f:
        print("Filename :", path, file=f)

orig_stdout = sys.stdout
f = open(fileName + ".csv", "w")
sys.stdout = f

print("RRD Name,Assigned Role")

sys.stdout = orig_stdout
f.close()

os.system("cat " + path + "draft.csv | tr 'b' ',' | tr -d \"'\" | cut -d , -f2,3 | tr -d ' ' | tr '_' ' ' >> " + path + fileName + ".csv")
os.system("rm " + path + "draft.csv")

print("\nSaving " + fileName + ".csv ...\n")
print(os.system("cat " + path + fileName + ".csv"))
os.system("start " + path + fileName + ".csv")
