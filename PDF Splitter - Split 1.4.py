#PDF Splitter and manager - Split
#Adding a way to "date stamp" all pages of pdf, better known as a small watermark, update 1.1
#Added error handling and success prompts that were previously missing, updated for new pypdf2 version


import PyPDF2 as pdf
from PyPDF2 import PdfMerger, PdfReader, PdfWriter
import tkinter as tk
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
import os
import sys
import glob
import time
from pathlib import Path
global rotation_angle
rotation_angle = 0

#root window
root=tk.Tk()
root.geometry("250x230")
root.title('Split')
root.configure()

L1=Label(root,text='', wraplength=250)
#last minute icon support
import icon
iconWindow=PhotoImage(data=icon.img)
root.tk.call('wm', 'iconphoto', root._w, iconWindow)


#begin program
pdfRead=''
pdfWrite=''
filePath=''
def pdfLoad():
    global pdfRead
    global filePath
    filePath=filedialog.askopenfilename()
    if not filePath: return # Stop if user cancels
    
    L1.config(text="Loaded PDF: "+os.path.basename(filePath))
    pdfRead=pdf.PdfReader(filePath)
    # no longer needed
    # pdfWrite=pdf.PdfWriter()

#------------RELOADING FILES-------------------
def pdfReload():
    global pdfRead
    if filePath: # Only try to reload if a file path exists
        pdfRead = PdfReader(filePath)
        L1.config(text="Loaded PDF: " + os.path.basename(filePath))

#------------SPLITTING FILES-------------------
def pdfSplitStart():
    if filePath == '':
        messagebox.showerror("Error", "Please load a PDF file first.")
        return
    global EnterSplit
    global topD
    topD=Toplevel()
    LtopD1=Label(topD, text="Please enter the page to split by.").pack()
    EnterSplit=Entry(topD, justify='center')
    BtopD1=Button(topD, text="Split File", command=pdfSplitContinue)

    EnterSplit.pack()
    BtopD1.pack()
    
def pdfSplitContinue():
    try:
        split_page_num = int(EnterSplit.get()) 
        total_pages = len(pdfRead.pages)

        if split_page_num <= 0 or split_page_num >= total_pages:
            messagebox.showerror("Error", f"Invalid split page. Please enter a number between 1 and {total_pages - 1}.")
            return

        writer1 = PdfWriter()
        writer2 = PdfWriter()

        for i in range(split_page_num):
            page = pdfRead.pages[i]
            writer1.add_page(page)

        for i in range(split_page_num, total_pages):
            page = pdfRead.pages[i]
            writer2.add_page(page)
        
        with open(str(os.path.dirname(filePath))+"/Part 1 "+str(os.path.basename(filePath)),'wb') as pdfExport1:
            writer1.write(pdfExport1)

        with open(str(os.path.dirname(filePath))+"/Part 2 "+str(os.path.basename(filePath)),'wb') as pdfExport2:
            writer2.write(pdfExport2)
        
        topD.destroy()
        messagebox.showinfo("Success", "PDF split successfully into two parts.")

    except ValueError:
        messagebox.showerror("Error", "Split page must be an integer.")
    

#-------------COMBINING 2 FILES------------------
def pdfCombineStart():
    if filePath == '':
        messagebox.showerror("Error", "Please load a PDF file first.")
        return
    global LtopC3
    topC=Toplevel()
    LtopC1=Label(topC, text="Your first file is "+os.path.basename(filePath)+".").pack()
    LtopC2=Label(topC, text="Please select another file to combine.").pack()
    BtopC1=Button(topC, text="File Select", command=CombineFile2)
    LtopC3=Label(topC, text='')
    BtopC2=Button(topC, text="Combine", command=pdfCombine)

    LtopC3.pack()
    BtopC1.pack()
    BtopC2.pack()

def CombineFile2():
    global cFile2
    cFile2=filedialog.askopenfilename()
    LtopC3.config(text="Combining: "+os.path.basename(filePath)+" and "+os.path.basename(cFile2)+".")
    
def pdfCombine():
    merger=PdfMerger()
    merger.append(filePath)
    merger.append(cFile2)
    pdfExport=open(str(os.path.dirname(filePath))+"/Combined Files "+str(os.path.basename(filePath)), 'wb')
    merger.write(pdfExport)
    messagebox.showinfo("Success", "PDFs combined successfully.")

#----------COMBINING MULTIPLE FILES------------
def multiCombineStart():
    topE=Toplevel()
    LtopE1=Label(topE, text="Your currently loaded file will be ignored for this process.").pack()
    LtopE2=Label(topE, text="The filename will be ""(first file name) (number of files) PDF Files.pdf")
    BtopE1=Button(topE, text="Please select multiple files using CTRL.",command=multiCombine)
    BtopE1.pack()
    LtopE2.pack()

def multiCombine():
    pdfList=filedialog.askopenfilenames()
    if not pdfList: return
    listCount=str(len(pdfList))
    merger=PdfMerger()
    for i in pdfList:
        merger.append(i)

    pdfExport=open(str(os.path.dirname(str(pdfList[0]))+"/"+os.path.basename(pdfList[0])+" "+listCount+" PDF Files.pdf"), 'wb')
    merger.write(pdfExport)
    messagebox.showinfo("Success", f"{listCount} PDFs combined successfully.")

#----------WATERMARKS AND DATE STAMPS----------
def dateStamp():
    if filePath == '':
        messagebox.showerror("Error", "Please load a PDF file first.")
        return
    global topF
    global LtopF2
    global LtopF3
    topF=Toplevel()
    topF.geometry("250x230")
    LtopF1=Label(topF, text="Watermark Creation")
    BtopF1=Button(topF, text="Select a PDF file of a watermark.", command=loadWmark, relief=GROOVE, borderwidth=2)
    LtopF2=Label(topF, text="")
    BtopF2=Button(topF, text="Click to watermark loaded PDF", command=processWmark, relief=GROOVE, borderwidth=2)
    LtopF3=Label(topF, text="")
    
    LtopF1.pack(pady=2)
    BtopF1.pack(pady=5, fill="both", expand=True)
    LtopF2.pack()
    BtopF2.pack()
    LtopF3.pack()
    
def loadWmark():
    global wmarkFile
    wmarkFile=filedialog.askopenfilename()
    topF.lift(aboveThis=root)
    LtopF2.config(text="Watermark Loaded: "+os.path.basename(wmarkFile))

def processWmark():
    LtopF3.config(text="Please allow a moment for processing...")
    topF.update_idletasks()
    
    wmark=PdfReader(wmarkFile)
    wmark_page=wmark.pages[0]

    pdf_sel=filePath
    pdf_tomark=PdfReader(pdf_sel)
    writer=PdfWriter()

    for page in pdf_tomark.pages:
        page.merge_page(wmark_page)
        writer.add_page(page)

    export_pdf=open(str(os.path.dirname(str(pdf_sel))+"/"+os.path.basename(pdf_sel)+" Watermarked.pdf"), 'wb')
    writer.write(export_pdf)

    LtopF3.config(text="Done!")
    messagebox.showinfo("Success", "Watermark applied successfully.")
    
#-------------ROTATING PAGE RANGE--------------
def pdfRotateB():
    if filePath == '':
        messagebox.showerror("Error", "Please load a PDF file first.")
        return
    if rotation_angle == 0:
        messagebox.showerror("Error", "Please select a rotation angle first.")
        return

    global EnterRotateB1, EnterRotateB2, topRangeB, LRangeB2
    topRangeB=Toplevel()
    LRangeB=Label(topRangeB, text="Please enter a range of pages.")
    LRangeB2=Label(topRangeB,text="")
    
    EnterRotateB1=Entry(topRangeB,justify='center')
    EnterRotateB2=Entry(topRangeB,justify='center')
    
    BRangeB=Button(topRangeB, text="Rotate Range", command=pdfRotateBContinue)

    LRangeB.pack()
    LRangeB2.pack()
    EnterRotateB1.pack()
    EnterRotateB2.pack()
    BRangeB.pack()
    

def pdfRotateBContinue():
    try:
        start_req = int(EnterRotateB1.get())
        end_req = int(EnterRotateB2.get())
        total_pages = len(pdfRead.pages)

        if end_req > total_pages or start_req < 1 or end_req < start_req:
             messagebox.showerror("Error","Please enter a valid range.")
             return

        LRangeB2.config(text="Rotating...")
        topRangeB.update_idletasks()

        writer = PdfWriter()
        
        for index, page in enumerate(pdfRead.pages):
            if (start_req - 1) <= index < end_req:
                page.rotate(rotation_angle)
            writer.add_page(page)

        output_filename = str(os.path.dirname(filePath))+"/Rotated Range "+str(start_req)+'-'+str(end_req)+" "+str(os.path.basename(filePath))
        with open(output_filename, 'wb') as pdfExport:
            writer.write(pdfExport)

        topRangeB.destroy()
        messagebox.showinfo("Success", "Page range rotated successfully.")
        pdfReload()

    except ValueError:
        messagebox.showerror("Error", "Page numbers must be integers.")

#-------------EXPORTING PAGE RANGE--------------
def pdfRangeA():
    if filePath == '':
        messagebox.showerror("Error", "Please load a PDF file first.")
        return
    global EnterRangeA1, EnterRangeA2, topRangeA, LRangeA2
    topRangeA=Toplevel()
    LRangeA=Label(topRangeA, text="Please enter a range of pages.")
    LRangeA2=Label(topRangeA,text="")
    
    EnterRangeA1=Entry(topRangeA,justify='center')
    EnterRangeA2=Entry(topRangeA,justify='center')
    
    BRangeA=Button(topRangeA, text="Export Range", command=pdfRangeAContinue)

    LRangeA.pack()
    LRangeA2.pack()
    EnterRangeA1.pack()
    EnterRangeA2.pack()
    BRangeA.pack()
    

def pdfRangeAContinue():
    try:
        start_req = int(EnterRangeA1.get())
        end_req = int(EnterRangeA2.get())
        total_pages = len(pdfRead.pages)

        if end_req > total_pages:
            messagebox.showerror("Error", f"Invalid page range. This PDF only has {total_pages} pages.")
            topRangeA.destroy()
            return

        if start_req < 1 or end_req < start_req:
             messagebox.showerror("Error","Please enter a valid range (e.g., from 2 to 5).")
             return

        LRangeA2.config(text=f"Exporting page {start_req} through page {end_req}")
        merger=PdfMerger()
        
        merger.append(filePath, pages=(start_req - 1, end_req))
        
        pdfExport=open(str(os.path.dirname(filePath))+"/Range "+str(start_req)+'-'+str(end_req)+" "+str(os.path.basename(filePath)), 'wb')
        merger.write(pdfExport)
        LRangeA2.config(text="Exported Successfully.")
        messagebox.showinfo("Success", "Page range exported successfully.")
        topRangeA.destroy()

    except ValueError:
        messagebox.showerror("Error", "Page numbers must be integers.")
        return
#----------TAKING OUT PAGES-----------------
def pdfYoinkerStart():
    if filePath == '':
        messagebox.showerror("Error", "Please load a PDF file first.")
        return
    global EnterYoink, topYoink
    topYoink=Toplevel()
    LYoink=Label(topYoink,text="Which page would you like to export?")
    EnterYoink=Entry(topYoink, justify='center')
    BYoink=Button(topYoink, text="Export Page", command=pdfYoinkerContinue)

    LYoink.pack()
    EnterYoink.pack()
    BYoink.pack()

def pdfYoinkerContinue():
    try:
        answer=EnterYoink.get()
        pageFind=int(answer)-1
        
        writer = PdfWriter()
        writer.add_page(pdfRead.pages[pageFind])
        
        pdfExport= open(str(os.path.dirname(filePath))+"/Export P"+str(pageFind+1)+' '+str(os.path.basename(filePath)), 'wb')
        writer.write(pdfExport)
        pdfExport.close()
        topYoink.destroy()
        messagebox.showinfo("Success", f"Page {pageFind + 1} exported successfully.")
    except (ValueError, IndexError):
        messagebox.showerror("Error", "Invalid page number.")


#---------- ROTATING PAGES------------------
def pdfRotateAll():
    if filePath == '':
        messagebox.showerror("Error", "Please load a PDF file first.")
        return
    if rotation_angle == 0:
        messagebox.showerror("Error", "Please select a rotation angle first.")
        return

    writer = PdfWriter()
    for page in pdfRead.pages:
        page.rotate(rotation_angle)
        writer.add_page(page)

    with open(str(os.path.dirname(filePath))+"/Rotated "+str(os.path.basename(filePath)), 'wb') as pdfExport:
        writer.write(pdfExport)
    messagebox.showinfo("Success", "Entire PDF rotated successfully.")
    pdfReload()

def pdfRotateStart():
    if filePath == '':
        messagebox.showerror("Error", "Please load a PDF file first.")
        return
    if rotation_angle == 0:
        messagebox.showerror("Error", "Please select a rotation angle first.")
        return

    global EnterRotate, topRotate
    topRotate=Toplevel()
    LRotate=Label(topRotate,text="Which page would you like to rotate?")
    EnterRotate=Entry(topRotate, justify='center')
    BRotate=Button(topRotate, text="Rotate Page",command=pdfRotateContinue)
    LRotate.pack()
    EnterRotate.pack()
    BRotate.pack()
    
    
def pdfRotateContinue():
    try:
        answer = EnterRotate.get()
        page_to_rotate = int(answer) - 1
        total_pages = len(pdfRead.pages)

        if not (0 <= page_to_rotate < total_pages):
            messagebox.showerror("Error", "Invalid page number.")
            return

        writer = PdfWriter()
        
        for index, page in enumerate(pdfRead.pages):
            if index == page_to_rotate:
                page.rotate(rotation_angle)
            writer.add_page(page)

        output_filename = str(os.path.dirname(filePath)) + "/Rotated P" + str(page_to_rotate + 1) + ' ' + str(os.path.basename(filePath))
        with open(output_filename, 'wb') as pdfExport:
            writer.write(pdfExport)

        topRotate.destroy()
        messagebox.showinfo("Success", f"Page {page_to_rotate + 1} rotated successfully.")
        pdfReload()

    except (ValueError, IndexError):
        messagebox.showerror("Error", "Invalid page number.")

B1=Button(root, text="Load PDF File", command=pdfLoad, relief=GROOVE, borderwidth=2)
B1.pack(fill="x",expand=True)
L1.pack(fill="both", expand=True)

#notebook windows
note=ttk.Notebook(root,height=200,width=250)
note.pack(fill="both", expand=True)
f1=ttk.Frame(note)
f2=ttk.Frame(note)
f3=ttk.Frame(note)
note.add(f1, text='Split')
note.add(f2, text='Combine')
note.add(f3, text='Rotate')

B2=Button(f1, text="Export Single Page", command=pdfYoinkerStart, relief=GROOVE, borderwidth=2)
B3=Button(f1, text="Export Page Range", command=pdfRangeA, relief=GROOVE, borderwidth=2)
B4=Button(f1, text="Split PDF", command=pdfSplitStart, relief=GROOVE, borderwidth=2)

B5=Button(f2, text="Combine PDFs", command=pdfCombineStart, relief=GROOVE, borderwidth=2)
B6=Button(f2, text="Combine Multiple PDFs", command=multiCombineStart, relief=GROOVE, borderwidth=2)
B7=Button(f2, text="Apply Watermarks", command=dateStamp, relief=GROOVE, borderwidth=2)

B8=Button(f3, text="Rotate PDF", command=pdfRotateAll, relief=GROOVE, borderwidth=2)
B9=Button(f3, text="Rotate Single Page", command=pdfRotateStart, relief=GROOVE, borderwidth=2)
B10=Button(f3, text="Rotate Page Range", command=pdfRotateB, relief=GROOVE, borderwidth=2)

B2.pack(pady=3)
B3.pack(pady=3)
B4.pack(pady=3)
B5.pack(pady=3)
B6.pack(pady=3)
B7.pack(pady=3)
B8.pack(pady=3)
B9.pack(pady=3)
B10.pack(pady=3)

def set_rotation_angle(angle):
    global rotation_angle
    rotation_angle = angle
    L9.config(text=f"Selected Angle: {rotation_angle}°")
#new rotate
L9=Label(f3, text="No Angle Selected.")
F9=ttk.Frame(f3)
B90=Button(F9, text="90°", command=lambda: set_rotation_angle(90))
B180=Button(F9, text="180°", command=lambda: set_rotation_angle(180))
B270=Button(F9, text="270°", command=lambda: set_rotation_angle(270))

L9.pack(pady=5)
F9.pack(pady=5)
B90.pack(side=LEFT, padx=5, expand=True)
B180.pack(side=LEFT, padx=5, expand=True)
B270.pack(side=LEFT, padx=5, expand=True)

root.mainloop()
