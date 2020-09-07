''' 
Script Name: run_itemizer.py
Author: Chris Lewis
Date: 04/10/2020
Purpose: The purpose of this script is to itemize Microsoft Office documents 
into their various component parts. This specific script creates an
instance of the GUI to run. The user may specify how they would like to
search for documents, the document types to search for, and the component
types to retrieve. They may also specify a directory to search for/save the
metadata and if they would like to keep the excess metdata that is stored in 
a temporary folder. After computation takes place a short summary of the files
generated is printed to the user in the GUI. Currently this script works for 
Word documents, Excel spreadsheets, and PowerPoint presentations.

Requirements:
- Running on macOS
- Python version 3.7.6 or higher
- All modules should be part of the Python standard library
'''

from tkinter import *
from itemizer_gui import ItemizerGUI


master = Tk()
application = ItemizerGUI(master)
master.mainloop()