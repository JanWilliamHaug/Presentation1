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
        print("Keywords found:", draft_keyword)
        #prints out "found tag:" whenever a tag is found

# function to display text when
# button is clicked
def clicked():
    lbl.configure(text = "Excel report has been generated")

    document = docx.Document('SRS_ACE_Pump_X01.docx')

    keyword1 = "PUMP"
    keyword2 = "ACE"
    keyword3 = "PRS"
    keyword4 = "SRS"
    name = "Keyword"
    title = "Tag"


    for paragraph in document.paragraphs:
        find_("PUMP", keyword1, paragraph)
        find_("ACE", keyword2, paragraph)
        find_("PRS", keyword3, paragraph)
        find_("SRS", keyword4, paragraph)

    excelFile = xw.Book()                #Creates an empty excel file
    excelFile.save('report.xlsx')        #Saves that excel file as "data1"

    ws1 = excelFile.sheets['Sheet1']
    ws1.range('A1').value = name         #Adds the string "Keyword" to A2
    ws1.range('B1').value = title        #Adds the string "Tag" to B1
    ws1.range('A2').value = keyword1     #Adds the first keyword to A2
    ws1.range('A3').value = keyword2     #Adds the second keyword to A3
    ws1.range('B2').value = keyword3     #Adds the third keyword to B2
    ws1.range('B3').value = keyword4     #Adds the fourth keyword to B3


    lbl2 = Label(root, text = "Keywords found:")
    lbl2.grid()
    lbl3 = Label(root, text = keyword1)
    lbl3.grid()
    lbl4 = Label(root, text = keyword2)
    lbl4.grid()
    lbl5 = Label(root, text = keyword3)
    lbl5.grid()
    lbl6 = Label(root, text = keyword4)
    lbl6.grid()


# button widget with red color text
btn = Button(root, text = "Generate report" ,
             fg = "red", command=clicked)
# set Button grid
btn.grid(column=1, row=0)

# Execute Tkinter
root.mainloop()








