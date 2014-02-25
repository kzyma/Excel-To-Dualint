#!/usr/bin/python
######################################################################
#
#   file: Excel-DualocIN.py
#   author: Ken Zyma
#
#   This file holds the main application for an Excel-Dualoc converter.
#   ...the input excel format can be found in the readme.txt
#
#   modification history:
#       Feb. 2014- created (author: Ken Zyma)
#
######################################################################

import Tkinter as Tk
import tkFileDialog,Tkconstants
import os
import ntpath
import ExcelToDualocConversion as EtoDConvert
import xlrd

class ExcelToDualocApp(Tk.Tk):
    
    def __init__(self,master):
        Tk.Tk.__init__(self,master)
        
        #empty StringVar, allows label to automatically update
        self.fileInputStr=Tk.StringVar()
        self.fileInput=""
        self.stdOut=Tk.StringVar()
        self.sheetInput=0
        self.sheetInputStr=Tk.StringVar()
        self.main=None
        
        self.initialize(master)

    def initialize(self,master):
        self.main=Tk.Frame(master,width=350, height=190)

        #returns filename for reading input
        getFileB=Tk.Button(master,text="File", command=self.directorySearchR, padx=40)
        getFileB.place(x=50,y=30)
        
        #options for opening the file
        self.file_opt = options = {}
        currentDirectory=os.getcwd()
        options['initialdir'] = currentDirectory
        options['title'] = 'Choose Excel File for Converting'

        fileName=Tk.Label(master, textvariable=self.fileInputStr)
        fileName.place(x=170,y=33)

        getSheetB=Tk.Button(master,text="Sheet",command=self.chooseSheet, padx=35)
        getSheetB.place(x=50,y=70)
        sheetNameL=Tk.Label(master,textvariable=self.sheetInputStr)
        sheetNameL.place(x=170,y=73)

        convertB=Tk.Button(master,text="Convert to DUALOC Input Format",
                           command=self.runConversion)
        convertB.place(x=50,y=110)

        stdMessageB=Tk.Label(master,textvariable=self.stdOut)
        stdMessageB.place(x=50,y=150)
        
        self.main.pack()

    def directorySearchR(self):
        
        def openExcelWorkbook(filePath):
            try:
                #do not need to explicity close file, xlrd handles this.
                f=xlrd.open_workbook(filePath)
            except IOError:
                print 'file',filePath,'failed to open'
                f=-1
                self.stdOut.set("Conversion Failed, file not found")
            except xlrd.XLRDError:
                print 'file',filePath,'is an unsupported format, see xlrd documentation'+\
                      ' https://github.com/python-excel/xlrd'
                f=-1
                self.stdOut.set("Unsupported format, see xlrd documentation")
            finally:
                return f

        #reset std output
        self.stdOut.set("")
        #returns opened file in read mode.
        fileName = tkFileDialog.askopenfilename(**self.file_opt)
        #get file name without path for display,
        #the next line is to handle both / and \ directory delimeters.
        if(fileName==""):
            return -1
        
        fileStr = ntpath.basename(fileName)
        
        #if fileStr doesnt fit on the screen...concatonate and append with ...
        fileStr=self.trim(fileStr,16)
            
        self.fileInputStr.set(fileStr)
        self.fileInput=fileName

        #set a default sheet
        b=openExcelWorkbook(self.fileInput)
        
        if (b==-1):
            return -1
        
        sheet0 = b.sheet_by_index(0)
        self.sheetInput=0
        sheetName=self.trim(sheet0.name,16)
        self.sheetInputStr.set(sheetName)
        return fileName

    def trim(self,string,size):
        if (len(string)>size):
            string=string[:(size-1)]
            string=string+"..."
            return string
        else:
            return string

    def chooseSheet(self):
        
        def setSheetAndExit(event):
            '''
            v=v.get()
            v=int(v)
            self.sheetInput=v
            '''            
            widget = event.widget
            selection= widget.curselection()
            v = widget.get(selection[0])
            
            #get sheet index
            sheetIndex=''
            j=0
            for i in excelWorkbook.sheets():
                if str(i.name)==v:
                    v=j
                j=j+1
                
            self.sheetInput=int(v)
            self.sheetInputStr.set(sheets[int(v)])
            sheetPopup.destroy()
            
        #reset std output
        self.stdOut.set("")
        #see if workbook uses multiple sheets
        excelWorkbook = xlrd.open_workbook(self.fileInput)
        sheets=[]
        
        for s in excelWorkbook.sheets():
            sheets.append(str(s.name))

        sheetPopup=Tk.Toplevel(self.master)
        
        #get x and y coordinets of root frame
        x = self.main.winfo_rootx()
        y = self.main.winfo_rooty()
   
        #make unresizable
        sheetPopup.resizable(width=None, height=None)
        sheetPopup.geometry('350x150+'+str(x)+'+'+str(y))
        sheetPopup.title("double-click to select")

        #add scrollbar, incase # of sheets exceeds max size
        scroller=Tk.Scrollbar(sheetPopup)
        scroller.pack(side=Tk.RIGHT,fill=Tk.Y)

        listbox=Tk.Listbox(sheetPopup)
        listbox.pack()
        
        #v=Tk.StringVar()
        for i in range(len(sheets)):
            '''
            b=Tk.Radiobutton(sheetPopup, text=sheets[i],
                             indicatoron=0,value=i,variable=v,
                             width=20,padx=20).pack(anchor=Tk.W)
            '''
            listbox.insert(Tk.END,sheets[i])
        '''
        choose=Tk.Button(sheetPopup,text="Choose Sheet", command=lambda: setSheetAndExit(v), padx=40)
        choose.pack()
        '''
        listbox.bind("<Double-Button-1>", setSheetAndExit)

        #attach listbox to scroller
        listbox.config(yscrollcommand=scroller.set)
        scroller.config(command=listbox.yview)


    def runConversion(self):
        #reset std output
        self.stdOut.set("")
        
        #run conversion with excel file input, return output filename (will be closed)
        if(self.fileInput==""):
            self.stdOut.set("Please enter a file")
            return -1
            
        convertFunctor = EtoDConvert.ExcelToDualocConversion()
        convertResult=convertFunctor(self.fileInput,self.sheetInput)
        
        if(convertResult==-1):
            self.stdOut.set("Conversion Failed, file not found")
        elif(convertResult==1):
            self.stdOut.set("Failure to open/create output file")
        elif(convertResult==3):
            self.stdOut.set("Incorrect excel file format")
        else:
            self.stdOut.set("Conversion Complete!")
        
if(__name__ == "__main__"):
    
    app = ExcelToDualocApp(None)
    app.title("Excel To DUALOC Input Converter")
    app.mainloop()


