# Generate Bulk Documents Python

It is a python script that generate bulk documents automatically from excel sheet data using replace concept. It can be used to generate bulk documents like letters, cards and certificates. It read data from an excel sheet and generate multiple documents.

## Demo
* Video clip on youtube of the script execution. https://youtu.be/SBubFoe84M0

## Important Note
* If this repository helped you to understand at least something new please give star this repository which motivates me to work further for the similar kinds for projects.

## Prerequisites

In order to run the python script, your system must have the following programs/packages installed.
* Python 3.8: Download it from https://www.python.org/downloads
* Pandas : Run in command prompt **pip install pandas**
* Xlrd : Run in command prompt **pip install xlrd**
* python-docx: Run in command prompt **pip install python-docx**

## Approach
* First need to clone this respiratory.
* Run python script which is in script.py using py script.py
* The script reads data from excel sheet, replace texts and generate documents.
* Loop execute until and unless all rows completes in excel sheet.

## Code
```
import pyautogui
import pandas
import datetime
import time
from docx import Document
import os

# Author @inforkgodara

# Read data from excel
excel_data = pandas.read_excel('data.xlsx', sheet_name='Recipient Details')
count = 0
directory = 'generated letters'

def replaceWord(oldString, newString, paragraph):
    if oldString in paragraph:
        inline = p.runs
        # Loop added to work with runs (strings with same style)
        for i in range(len(inline)):
            if oldString in inline[i].text:
                text = inline[i].text.replace(oldString, newString)
                inline[i].text = text

# Iterate excel rows till to finish
for column in excel_data['Recipient Full Name'].tolist():
    document = Document('letter template.docx')
    doc = document
    empName = excel_data['Recipient Full Name'][count]
    for p in doc.paragraphs:
        replaceWord('RECIPIENT NAME', excel_data['Recipient Full Name'][count], p.text)
        replaceWord('FIRST NAME', excel_data['Recipient First Name'][count], p.text)
        replaceWord('TITLE', excel_data['Recipient Title'][count], p.text)
        replaceWord('COMPANY NAME', excel_data['Recipient Company Name'][count], p.text)
        replaceWord('STREET ADDRESS', excel_data['Recipient Street Address'][count], p.text)
        replaceWord('CITY, ST ZIP CODE', str(excel_data['Recipient City, ST ZIP Code'][count]), p.text)

    try:
        path = os.getcwd()+"/"+directory+"/"+empName
        os.mkdir(path)
    except OSError:
        a = 10
    doc.save(os.getcwd()+"/"+directory+"/"+empName+"/"+empName+' Latter.docx')
    print("Letter generated for " + empName)
    count = count + 1

print("Total letters are created " + str(count))
```
Video on youtube: https://youtu.be/SBubFoe84M0
