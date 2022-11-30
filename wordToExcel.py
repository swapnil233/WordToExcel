'''
Instructions for future users:
1. Download and install the right version of pywin32 from https://github.com/mhammond/pywin32/releases 
2. Download the tagged word docs, and rename them to "Program - Person.docx"
3. Place all those docs in your Documents folder
4. Run this script
'''

# For the Word document
import win32com.client as win32
from win32com.client import constants
word = win32.gencache.EnsureDispatch('Word.Application')
word.Visible = False
filepath = "Envirosoft.docx"

import re

# For the Excel document
from openpyxl import Workbook
from openpyxl.styles import Font
workbook = Workbook()
worksheet = workbook.active

# Look for all the .docx files in the C:\Users\IqbalH\Documents folder and list their names in an array
import os
path = "C:\\Users\\IqbalH\\Documents"

# List of word documents in the Documents folder
wordDocs = []
for file in os.listdir(path):
    if file.endswith(".doc") or file.endswith(".docx"):
        wordDocs.append(file)

tags = []
descriptions = []

def get_comments(filepath):
    doc = word.Documents.Open(filepath) 
    doc.Activate()
    activeDoc = word.ActiveDocument
    
    for comment in activeDoc.Comments:
        #checking if this is a top-level comment 
        if comment.Ancestor is None:
            
            # Author of the comment
            #print("Comment by: " + comment.Author)
            
            # text of the comment
            #print("Comment text: " + comment.Range.Text) 
            
            # text of the original document where the comment is anchored
            #print("Regarding: " + comment.Scope.Text)
            
            # Insert the text and comment into the arrays. They should both be the same length
            tags.append(comment.Range.Text.replace("\n", "").strip())
            descriptions.append(comment.Scope.Text.replace("\n", "").strip())
            
    doc.Close()

# Clean array of illegal characters
def clean_arr(tags, descriptions):
    cleaned_tags = []
    cleaned_descriptions = []
    
    for i in range(len(tags)):
        
        # Tags (the actual comments by us)
        cleaned_tags.append(tags[i].replace("\x05", "").replace("\r", "").replace("\r\r", "").replace("\x0b", ", "))
        
        # Tag descriptions
        cleaned_descriptions.append(descriptions[i]
                                  .replace("\x05", "")
                                  .replace("\r", "")
                                  .replace("\r\r", "")
                                  .replace("\x0b", " ")
                                  .replace("[Ron] ", "")
                                  .replace("[user] ", "")
                                  .replace("[speaker] ", "")
                                  .replace("[", "")
                                  .replace("]", ""))
        
    for i in range(len(cleaned_descriptions)):
        # Removing timestamps of the form [00:00]
        if re.search(r"\d\d:\d\d", cleaned_descriptions[i]):
            cleaned_descriptions[i] = re.sub(r"\d\d:\d\d", "", cleaned_descriptions[i])
        
    return cleaned_tags, cleaned_descriptions

# Write the arrays to the Excel document in columns
'''
id | summary | description | program | tag | person | comments | technical name | source
'''

# Headers
worksheet.append(["ID", "Summary", "Description", "Program", "Tag", "Person", "Our comments/questions", "Technical name", "Source"])

# Bold headers
for cell in worksheet["1:1"]:
    cell.font = Font(color='000000', bold=True)

# Make excel file if it doesn't exist 
if not os.path.exists("Full.xlsx"):
    workbook.save("Full.xlsx")


def write_to_excel(cleaned_tags : list, cleaned_descriptions : list, program : str, person : str):
    
    """
    Adds the rows of information to the "Full.xlsx" file with the following columns:
    id | summary | description | program | tag | person | comments | technical name | source

    Args:
        cleaned_tags: An array of strings, each string is a comment
        cleaned_descriptions: An array of strings, each string is the text of the comment's scope (the text that the comment is about)
        program: The name of the program, eg "Envirosoft", taken from the filename
        person: The name of the person, eg "Ashley Mathew", taken from the filename
    
    Returns:
        None
    
    Raises:
        None
    """
    
    for i in range(len(cleaned_tags)):
        
        current_id = "L" + str(i+1)
        summary = ""
        description = cleaned_descriptions[i]
        tag = cleaned_tags[i]
        comments = ""
        technical_name = ""
        source = ""
        
        # Append the row to the excel file
        worksheet.append([current_id, summary, description, program, tag, person, comments, technical_name, source])
        
    workbook.save("Full.xlsx")

# Loop through each word document in the "wordDocs" array, get the comments, clean the arrays, and write to Excel
for i in range(len(wordDocs)):
    get_comments(wordDocs[i])
    cleaned_tags, cleaned_descriptions = clean_arr(tags, descriptions)
    
    # Split the word document name into the program and person
    program = wordDocs[i].split(" - ")[0]
    person = wordDocs[i].split(" - ")[1].split(".")[0]
    
    write_to_excel(cleaned_tags, cleaned_descriptions, program, person)
    
    program = ""
    person = ""
    tags = []
    descriptions = []