#!/usr/bin/env python
# coding: utf-8

# In[1]:


## Import packages 

from PyPDF2 import PdfReader
import os 
import re 
import shutil
import sys
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo


# In[9]:


### Functions 
def RunApp():

    
    root = tk.Tk()
    root.title('PDF Renamer')
    root.resizable(False, False)
    root.geometry('400x200')
    
    # open button
    open_button = ttk.Button(
        root,
        text = 'Open Files',
        command = RenamePDF)

    # close button
    close_button = ttk.Button(
        root, 
        text = 'Close', 
        command=root.destroy)

    open_button.pack(expand = False)
    close_button.pack(expand = False)

    # Bring window to front of users screen
    root.lift()
    root.attributes('-topmost', True)
    root.after_idle(root.attributes, '-topmost', False)
    root.mainloop()


def select_files():
    
    file_chosen = tk.messagebox.askokcancel(title = "Choose files to rename", message = "Please select the files you wish you rename!")
    
  
    
    filenames = fd.askopenfilenames(
        title='Open files',
        initialdir='/',
        filetypes=(("PDF Files", "*.pdf"),))
    
    if not filenames: 
        tk.messagebox.showwarning(title = "No files were selected!", message = "No files were selected. Aborting application!")
    
    return(filenames)

def select_folder():
    """
    Function used for user to select the final folder to store renamed PDF files.
    """ 
    tk.messagebox.askokcancel(title = "Choose folder to save renamed files", message = "Please select the folder"                                                                                         " you wish to save the renamed files to")
    select_folder = fd.askdirectory()
    return(select_folder)

def RenamePDF():
    
    # User to select files to rename and folder to output the renamed files
    filenames = select_files()
    New_Folder = select_folder()
    
    # Loop through each file selected, find the pattern required then copy original file 
    # and rename using the required pattern. 
    for file in filenames:
        reader = PdfReader(file)
        text = ""
    
        for page in reader.pages:
            text += page.extract_text() + "\n"
            
            # Retreiving the CLient Code and Invoice number 
            ClientCode = re.search('(?<=Comments:).*', text).group(0).strip()
            InvNo = re.search('(?<=Invoice\sNo\s:\s)(\d+)', text).group(0).strip()
            
            # Creating the new PDF path
            NewPDFPath = New_Folder + "/" + ClientCode + "_" + InvNo + ".pdf"

            # Copy existing files and replace into selected folder 
            shutil.copyfile(file, NewPDFPath)
            
    tk.messagebox.showinfo(title = "Complete!", message = str(len(filenames)) + " files were copied to " + New_Folder)


# In[10]:


RunApp()


# In[ ]:




