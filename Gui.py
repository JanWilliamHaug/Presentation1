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

# function to display text when
# button is clicked
def clicked():
    lbl.configure(text = "Report has been generated")
    excelFile = xw.Book()                #Creates an empty excel file
    excelFile.save('report.xlsx')        #Saves that excel file as "data1"

    ws1 = excelFile.sheets['Sheet1']



# button widget with red color text
# inside
btn = Button(root, text = "Generate report" ,
             fg = "red", command=clicked)
# set Button grid
btn.grid(column=1, row=0)

# Execute Tkinter
root.mainloop()


