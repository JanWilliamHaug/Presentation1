#1. remember to install "python-docx" and "xlwings" first
#2. Import libraries
import xlwings as xw
import docx
from docx import Document
from tkinter import *

# create root window
root = Tk()

# root window title and dimension
root.title("TARGEST")
# Set geometry(widthxheight)
root.geometry('350x200')

# adding a label to the root window
lbl = Label(root, text = "Do you want to generate an excel report?")
lbl.grid()

def find_(paragraph_keyword, draft_keyword, paragraph):
    if paragraph_keyword in paragraph.text:
        print("found tag:", draft_keyword)
        #prints out "found tag:" whenever a tag is found

# function to display text when
# button is clicked
def clicked():
    lbl.configure(text = "Report has been generated")

    document = docx.Document('test10.docx')
    print(document.paragraphs[0].text)
    print(document.paragraphs[1].text)

    product = "TARGEST"
    coFounder1 = "Jan"
    coFounder2 = "Adrian"
    coFounder3 = "Stephania"
    name = "NAME"
    title = "TITLE"
    title2 = "Co-Founder"

    for paragraph in document.paragraphs:
        find_("product", product, paragraph)
        find_("Co-Founder1", coFounder1, paragraph)
        find_("Co-Founder2", coFounder2, paragraph)
        find_("Co-Founder3", coFounder3, paragraph)

    excelFile = xw.Book()                #Creates an empty excel file
    excelFile.save('report.xlsx')        #Saves that excel file as "data1"

    ws1 = excelFile.sheets['Sheet1']
    ws1.range('A1').value = name         #Adds the string "Name" to A2
    ws1.range('B1').value = title        #Adds the string "Title" to B1
    ws1.range('A2').value = coFounder1   #Adds name of Co-Founder 1 to A2
    ws1.range('A3').value = coFounder2   #Adds name of Co-Founder 2 to A3
    ws1.range('A4').value = coFounder3   #Adds name of Co-Founder 3 to A3
    ws1.range('C1').value = product      #Adds name of the product to C1
    ws1.range('B2:B4').value = title2    #Adds  titles






# button widget with red color text
# inside
btn = Button(root, text = "Generate report" ,
             fg = "red", command=clicked)
# set Button grid
btn.grid(column=1, row=0)

# Execute Tkinter
root.mainloop()








