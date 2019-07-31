#PDF Splitter and manager - Split
#Adding a way to "date stamp" all pages of pdf, better known as a small watermark, update 1.1


import PyPDF2 as pdf
from PyPDF2 import PdfFileMerger, PageRange, PdfFileReader, PdfFileWriter
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
from tqdm import tqdm
from tqdm import tqdm_gui


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
    global pdfWrite
    global filePath
    filePath=filedialog.askopenfilename()
    
    L1.config(text="Loaded PDF: "+os.path.basename(filePath))
    pdfRead=pdf.PdfFileReader(filePath)
    pdfWrite=pdf.PdfFileWriter()

#------------SPLITTING FILES-------------------
def pdfSplitStart():
    global EnterSplit
    topD=Toplevel()
    LtopD1=Label(topD, text="Please enter the page to split by.").pack()
    EnterSplit=Entry(topD, justify='center')
    BtopD1=Button(topD, text="Split File", command=pdfSplitContinue)

    EnterSplit.pack()
    BtopD1.pack()
    
def pdfSplitContinue():
    halfwayP=int(EnterSplit.get())
    print(halfwayP)
    pdfWrite2=pdf.PdfFileWriter()

    xStart=0
    xEnd=pdfRead.getNumPages()
    pageList1=list(range(xStart,halfwayP))
    pageList2=list(range(halfwayP, xEnd))
    
    #add and export all pages up to the user specified, then use user specified to lastpage
    for i in pageList1:
        pageA=pdfRead.getPage(i)
        pdfWrite.addPage(pageA)
    pdfExport=open(str(os.path.dirname(filePath))+"/Part 1 "+str(os.path.basename(filePath)),'wb')
    pdfWrite.write(pdfExport)
    
        
    for i in pageList2:
        pageB=pdfRead.getPage(i)
        pdfWrite2.addPage(pageB)
    pdfExport2=open(str(os.path.dirname(filePath))+"/Part 2 "+str(os.path.basename(filePath)),'wb')
    pdfWrite2.write(pdfExport2)
    
    

#-------------COMBINING 2 FILES------------------
def pdfCombineStart():
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
    merger=PdfFileMerger()
    merger.append(pdfRead, 'rb')
    merger.append(cFile2, 'rb')
    pdfExport=open(str(os.path.dirname(filePath))+"/Combined Files "+str(os.path.basename(filePath)), 'wb')
    merger.write(pdfExport)

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
    print(pdfList[1])
    listCount=str(len(pdfList))
    merger=PdfFileMerger()
    for i in pdfList:
        pdfListRead=pdf.PdfFileReader(i)
        merger.append(i)

    pdfExport=open(str(os.path.dirname(str(pdfList[0]))+"/"+os.path.basename(pdfList[1])+" "+listCount+" PDF Files.pdf"), 'wb')
    merger.write(pdfExport)

#----------WATERMARKS AND DATE STAMPS----------
def dateStamp():
    topF=Toplevel()
    LtopF1=Label(topF, text="Watermark Creation")
    BtopF1=Button(topF, text="Select a PDF watermark.", command=loadWmark)
    BtopF2=Button(topF, text="Select a PDF to Watermark", command=processWmark)
    LtopF2=Label(topF, text="") #padding
    LtopF1.pack(pady=2)
    BtopF1.pack(pady=5)
    BtopF2.pack()
    LtopF2.pack(pady=3)

def loadWmark():
    global wmarkFile
    wmarkFile=filedialog.askopenfilename()

#error - pdffilereader not defined
def processWmark():
    wmark=PdfFileReader(wmarkFile)
    wmark_page=wmark.getPage(0)

    pdf_sel=filedialog.askopenfilename()
    pdf_tomark=PdfFileReader(pdf_sel)
    writer=PdfFileWriter()

    for page in tqdm_gui(range(pdf_tomark.getNumPages())):
        pdf_tomark_page=pdf_tomark.getPage(page)
        pdf_tomark_page.mergePage(wmark_page)
        writer.addPage(pdf_tomark_page)

    export_pdf=open(str(os.path.dirname(str(pdf_sel))+"/"+os.path.basename(pdf_sel)+" Watermarked.pdf"), 'wb')
    writer.write(export_pdf)

    
    
#-------------ROTATING PAGE RANGE--------------
def pdfRotateB():
    E1ans=E1.get()
    E1ans_int=int(E1ans)
    global EnterRotateB1
    global EnterRotateB2
    global topRangeB
    global LRangeB2
    print(" ")
    topRangeB=Toplevel()
    LRangeB=Label(topRangeB, text="Please enter a range of pages.")
    LRangeB2=Label(topRangeB,text="")
    
    EnterRotateB1=Entry(topRangeB,justify='center')
    
    
    EnterRotateB2=Entry(topRangeB,justify='center')
    
    #to make sure the pages aren't entered incorrectly, subtract A2 from A1, if negative then throw an error
    BRangeB=Button(topRangeB, text="Rotate Range", command=pdfRotateBContinue)

    LRangeB.pack()
    LRangeB2.pack()
    EnterRotateB1.pack()
    EnterRotateB2.pack()
    BRangeB.pack()
    

def pdfRotateBContinue():
    PDFRotate1=EnterRotateB1.get()
    PDFRotate2=EnterRotateB2.get()
    print(PDFRotate1)
    print(PDFRotate2)
    print(int(PDFRotate2)-int(PDFRotate1))
    
    if int(PDFRotate2)-int(PDFRotate1) <= -0.1:
        topRangeB.destroy()

        messagebox.showerror("Error","Please enter a valid range.")

        pdfRotateB()

    else:
        LRangeB2.config(text="Rotating page "+PDFRotate1+" through page "+PDFRotate2)
        merger=PdfFileMerger()
        pdfPageRange=str(int(PDFRotate1)-1)+":"+str(int(PDFRotate2))

        xRotate=int(PDFRotate1)-1
        yRotate=int(PDFRotate2)
        listNum=list(range(xRotate,yRotate))

        #iterate over list of pages and rotate each one
        for i in listNum:
            page=pdfRead.getPage(i)
            page.rotateClockwise(E1ans_int)
            pdfWrite.addPage(page)
        
        pdfExport=open(str(os.path.dirname(filePath))+"/Rotated Range "+str(PDFRotate1)+'-'+str(PDFRotate2)+" "+str(os.path.basename(filePath)), 'wb')
        
        pdfWrite.write(pdfExport)
        LRangeB2.config(text="Rotated Successfully.")
        topRangeB.destroy()
        pdfRotateB()

#-------------EXPORTING PAGE RANGE--------------
def pdfRangeA():
    global EnterRangeA1
    global EnterRangeA2
    global topRangeA
    global LRangeA2
    print(" ")
    topRangeA=Toplevel()
    LRangeA=Label(topRangeA, text="Please enter a range of pages.")
    LRangeA2=Label(topRangeA,text="")
    
    EnterRangeA1=Entry(topRangeA,justify='center')
    
    
    EnterRangeA2=Entry(topRangeA,justify='center')
    
    #to make sure the pages aren't entered incorrectly, subtract A2 from A1, if negative then throw an error
    BRangeA=Button(topRangeA, text="Export Range", command=pdfRangeAContinue)

    LRangeA.pack()
    LRangeA2.pack()
    EnterRangeA1.pack()
    EnterRangeA2.pack()
    BRangeA.pack()
    

def pdfRangeAContinue():
    PDFRange1=EnterRangeA1.get()
    PDFRange2=EnterRangeA2.get()
    print(PDFRange1)
    print(PDFRange2)
    print(int(PDFRange2)-int(PDFRange1))
    
    if int(PDFRange2)-int(PDFRange1) <= -0.1:
        topRangeA.destroy()

        messagebox.showerror("Error","Please enter a valid range.")

        pdfRangeA()

    else:
        LRangeA2.config(text="Exporting page "+PDFRange1+" through page "+PDFRange2)
        #exports, but misses a page.
        merger=PdfFileMerger()
        #-1 for pdfrange1 for proper list reference (starts from 0)
        pdfPageRange=str(int(PDFRange1)-1)+":"+str(int(PDFRange2))
        merger.append(pdfRead, pages=PageRange(pdfPageRange))
        pdfExport=open(str(os.path.dirname(filePath))+"/Range "+str(PDFRange1)+'-'+str(PDFRange2)+" "+str(os.path.basename(filePath)), 'wb')
        merger.write(pdfExport)
        LRangeA2.config(text="Exported Successfully.")

#----------TAKING OUT PAGES-----------------
def pdfYoinkerStart():
    #yoinks out a page
    global EnterYoink
    global topYoink
    topYoink=Toplevel()
    LYoink=Label(topYoink,text="Which page would you like to export?")
    EnterYoink=Entry(topYoink, justify='center')
    BYoink=Button(topYoink, text="Export Page", command=pdfYoinkerContinue)

    LYoink.pack()
    EnterYoink.pack()
    BYoink.pack()

def pdfYoinkerContinue():
    answer=EnterYoink.get()
    pageFind=int(answer)-1
    print(pageFind)

    page=pdfRead.getPage(pageFind)
    pdfWrite.addPage(page)
        
    pdfExport= open(str(os.path.dirname(filePath))+"/Export P"+str(pageFind+1)+' '+str(os.path.basename(filePath)), 'wb')
    pdfWrite.write(pdfExport)
    pdfExport.close()
    topYoink.destroy()


#---------- ROTATING PAGES------------------
def pdfRotateAll():
    #rotate all pages in pdf
    E1ans=E1.get()
    E1ans_int=int(E1ans)
    for pageNum in range(pdfRead.numPages):
        page = pdfRead.getPage(pageNum)
        page.rotateClockwise(E1ans_int)
        pdfWrite.addPage(page)

    
    pdfExport= open(str(os.path.dirname(filePath))+"/Rotated "+str(os.path.basename(filePath)), 'wb')
    pdfWrite.write(pdfExport)
    pdfExport.close()
  
def pdfRotateStart():
    #Creates popup interface for page specification
    global EnterRotate
    global topRotate
    topRotate=Toplevel()
    LRotate=Label(topRotate,text="Which page would you like to rotate?")
    EnterRotate=Entry(topRotate, justify='center')
    #top.bind('<Return>',pdfRotateContinue)
    BRotate=Button(topRotate, text="Export Page",command=pdfRotateContinue)
    LRotate.pack()
    EnterRotate.pack()
    BRotate.pack()
    
    
def pdfRotateContinue():
    E1ans=E1.get()
    E1ans_int=int(E1ans)
    answer=EnterRotate.get()
    pageFind=int(answer)-1
    print(pageFind)
    page=pdfRead.getPage(pageFind)
    page.rotateClockwise(E1ans_int)
    pdfWrite.addPage(page)
        
    pdfExport= open(str(os.path.dirname(filePath))+"/Rotated P"+str(pageFind+1)+' '+str(os.path.basename(filePath)), 'wb')
    pdfWrite.write(pdfExport)
    pdfExport.close()
    topRotate.destroy()
    

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


#entry support for rotate
global E1ans_pre
E1=Entry(f3, justify='center')
L2=Label(f3, text="Please enter a degree to rotate clockwise by.")
E1ans_pre=E1.get()




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

L2.pack()
E1.pack()


root.mainloop()
