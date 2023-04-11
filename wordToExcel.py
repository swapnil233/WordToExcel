'''
Instructions for future users:
1. Download and install the appropriate version of pywin32 from https://github.com/mhammond/pywin32/releases
2. Download the tagged Word documents, and rename them to "Program - Person.docx"
3. Place all these documents in your Documents folder
4. Run this script
'''

import os
import re
import win32com.client as win32
from win32com.client import constants
from openpyxl import Workbook
from openpyxl.styles import Font

# Initialize Word and Excel components
word = win32.gencache.EnsureDispatch('Word.Application')
word.Visible = False
workbook = Workbook()
worksheet = workbook.active

# Set the directory path for the Documents folder
path = os.path.expanduser('~/Documents')

# Retrieve all Word documents in the Documents folder
word_docs = [file for file in os.listdir(path) if file.endswith(('.doc', '.docx'))]

tags = []
descriptions = []

def get_comments(filepath):
    doc = word.Documents.Open(filepath)
    doc.Activate()
    active_doc = word.ActiveDocument
    
    for comment in active_doc.Comments:
        if comment.Ancestor is None:
            tags.append(comment.Range.Text.replace("\n", "").strip())
            descriptions.append(comment.Scope.Text.replace("\n", "").strip())
            
    doc.Close()

def clean_data(tags, descriptions):
    cleaned_tags = [tag.replace("\x05", "").replace("\r", "").replace("\r\r", "").replace("\x0b", ", ") for tag in tags]
    
    cleaned_descriptions = []
    for description in descriptions:
        cleaned_description = (description.replace("\x05", "")
                                         .replace("\r", "")
                                         .replace("\r\r", "")
                                         .replace("\x0b", " ")
                                         .replace("[Ron] ", "")
                                         .replace("[user] ", "")
                                         .replace("[speaker] ", "")
                                         .replace("[", "")
                                         .replace("]", ""))
        cleaned_descriptions.append(re.sub(r"\d\d:\d\d", "", cleaned_description))
    
    return cleaned_tags, cleaned_descriptions

def write_to_excel(cleaned_tags, cleaned_descriptions, program, person):
    for i in range(len(cleaned_tags)):
        current_id = "L" + str(i+1)
        summary = ""
        description = cleaned_descriptions[i]
        tag = cleaned_tags[i]
        comments = ""
        technical_name = ""
        source = ""
        
        worksheet.append([current_id, summary, description, program, tag, person, comments, technical_name, source])
        
    workbook.save("Full.xlsx")

# Create Excel file if it doesn't exist
if not os.path.exists("Full.xlsx"):
    worksheet.append(["ID", "Summary", "Description", "Program", "Tag", "Person", "Our comments/questions", "Technical name", "Source"])
    for cell in worksheet["1:1"]:
        cell.font = Font(color='000000', bold=True)
    workbook.save("Full.xlsx")

# Process all Word documents and save comments in Excel
for doc in word_docs:
    get_comments(doc)
    cleaned_tags, cleaned_descriptions = clean_data(tags, descriptions)
    program, person = doc.split(" - ")[0], doc.split(" - ")[1].split(".")[0]
    write_to_excel(cleaned_tags, cleaned_descriptions, program, person)
    tags.clear()
    descriptions.clear()