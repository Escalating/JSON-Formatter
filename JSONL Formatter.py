#Copyright 2018 Dillon Jennings

# This program is free software; you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation; either version 2 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program; if not, see <http://www.gnu.org/licenses/>.

import tkinter as tk
import re, xlsxwriter, os
from tkinter import *
from tkinter import Tk
from tkinter.filedialog import askopenfilename, asksaveasfilename
from tkinter.messagebox import askyesno, showerror, showinfo

window = Tk()

def openFile():
    try:
        ftypes = [('JL file',"*.jl"), ('JSON file', "*.json")]
        ttl = "Title"
        dir1 = 'C:\\'
        window.fileName = askopenfilename(filetypes = ftypes, initialdir = dir1, title = ttl)
        print(window.fileName)
        selectedFile = window.fileName
        top = Toplevel(window)
        with open(selectedFile, "r") as f:
            global fileText
            fileText = f.read()
            def parseFile():
                global fileText
                reg = r"\[\"(.+?)\"\]|[\[\]]+"
                parsed = Toplevel(window)
                parsed.maxsize(720, 720)
                parsed.title(selectedFile)
                scrollbar = Scrollbar(parsed, orient=VERTICAL)
                scrollbar.pack(side=RIGHT, fill=Y, expand=FALSE)
                scrollbar2 = Scrollbar(parsed, orient=HORIZONTAL)
                scrollbar2.pack(side=BOTTOM, fill=X, expand=FALSE)
                display = Text(parsed, xscrollcommand=scrollbar2.set, yscrollcommand=scrollbar.set, wrap=tk.NONE, width=300, height=300)
                display.pack(side=TOP, fill=BOTH, expand=TRUE)
                scrollbar.config(command=display.yview)
                scrollbar2.config(command=display.xview)
                matches = re.findall(reg, fileText)
                top.destroy()
                count = 0
                for match in matches:
                    text1 = re.sub(r"\"\, \"\, \"\, \"", ', ', match)
                    text2 = re.sub(r"\"\, \" \"\, \"", ' ', text1)
                    text3 = re.sub(r"\"\, \"", ' ', text2)
                    fileText = text3
                    display.insert(END, fileText + '\n')
                    
                def saveFile():
                    try:
                        file_options = options = {}
                        options['filetypes'] = [('All types', '*.*'), ('Normal text file', '*.txt')]
                        options['initialfile'] = 'parsedfile.txt'
                        options['parent'] = parsed
                        save_file = asksaveasfilename(**file_options)
                        data = display.get('1.0', END)[:-1]
                        f = open(save_file, 'w')
                        f.write(data)
                        parsed.destroy()
                        
                    except FileNotFoundError:
                        print("File was not selected, or not found.")
                        parsed.destroy()
                
                save = askyesno("Save", "Would you like to save this file?", parent=parsed)
                if save == True:
                    saveFile()
                else:
                    parsed.destroy()
                    
            top.maxsize(720, 720)
            top.title(selectedFile)
            scrollbar = Scrollbar(top, orient=VERTICAL)
            scrollbar.pack(side=RIGHT, fill=Y, expand=FALSE)
            scrollbar2 = Scrollbar(top, orient=HORIZONTAL)
            scrollbar2.pack(side=BOTTOM, fill=X, expand=FALSE)
            display = Text(top, xscrollcommand=scrollbar2.set, yscrollcommand=scrollbar.set, wrap=tk.NONE, width=300, height=300)
            display.pack(side=TOP, fill=BOTH, expand=TRUE)
            display.insert('1.0', fileText)
            scrollbar.config(command=display.yview)
            scrollbar2.config(command=display.xview)

            answer = askyesno("Parse", "Would you like to parse this file?", parent=top)
            if answer == True:
                parseFile()
            else:
                top.destroy()

    except FileNotFoundError:
        print("File was not selected, or not found.")
        top.destroy()
        
def selectFileToFormat():
    try:
        ftypes = [('Normal text file',"*.txt")]
        ttl = "Title"
        dir1 = 'C:\\'
        window.fileName = askopenfilename(filetypes = ftypes, initialdir = dir1, title = ttl)
        print(window.fileName)
        selectedFile = window.fileName
        top = Toplevel(window)
        with open(selectedFile, "r") as f:
            fileText = f.read()
            lines = fileText.splitlines()
            print(len(lines))
            top.maxsize(720, 720)
            top.title(selectedFile)
            scrollbar = Scrollbar(top, orient=VERTICAL)
            scrollbar.pack(side=RIGHT, fill=Y, expand=FALSE)
            scrollbar2 = Scrollbar(top, orient=HORIZONTAL)
            scrollbar2.pack(side=BOTTOM, fill=X, expand=FALSE)
            display = Text(top, xscrollcommand=scrollbar2.set, yscrollcommand=scrollbar.set, wrap=tk.NONE, width=300, height=300)
            display.pack(side=TOP, fill=BOTH, expand=TRUE)
            display.insert('1.0', fileText)
            scrollbar.config(command=display.yview)
            scrollbar2.config(command=display.xview)

            def formatFile():
                top2 = Toplevel(top)
                top2.resizable(0, 0)
                varcols = IntVar()
                l1 = Label(top2, text="How many columns? (Max of 5)")
                l1.pack()
                entry = Entry(top2, textvariable=varcols)
                entry.pack()
                def storecols():
                    global on, tw, th, fo, fi, cols
                    on = False
                    tw = False
                    th = False
                    fo = False
                    fi = False
                    cols = []
                    numcols = varcols.get()
                    
                    if numcols == 0:
                        showerror("Error", "You have to have at least 1 column.", parent=top2)
                        top2.destroy()
                        formatFile()
                    elif numcols == 1:
                        cols = [1]
                        on = True
                    elif numcols == 2:
                        cols =[1,2]
                        tw = True
                    elif numcols == 3:
                        cols = [1,2,3]
                        th = True
                    elif numcols == 4:
                        cols = [1,2,3,4]
                        fo = True
                    elif numcols == 5:
                        cols = [1,2,3,4,5]
                        fi = True
                    elif numcols > 5:
                        showerror("Error", "Max amount of columns is 5.", parent=top2)
                    
                    if on == True or tw == True or th == True or fo == True or fi == True:
                        print(numcols)
                        maxcols = [1,2,3,4,5]
                        top2.destroy()
                        top3 = Toplevel(top)
                        top3.resizable(0, 0)
                        varcolnames = StringVar()
                        l2 = Label(top3, text="Enter the names of the columns.")
                        l2.pack()
                        name_list = []
                        for col in cols:
                            row = Frame(top3)
                            lab = Label(row, width=15, text=col, anchor='w')
                            ent = Entry(row)
                            row.pack(side=TOP, fill=X, padx=5, pady=5)
                            lab.pack(side=LEFT)
                            ent.pack(side=RIGHT, expand=YES, fill=X)
                            name_list.append(ent)

                        def makeSheet():
                            nl = []
                            for name in name_list:
                                print(name.get())
                                name = name.get()
                                nl.append(name)
                                
                            top3.destroy()
                            print(nl)
                            file = os.path.splitext(os.path.basename(selectedFile))[0]
                            workbook = xlsxwriter.Workbook('%s.xlsx' % (file))
                            worksheet = workbook.add_worksheet()
                            bold = workbook.add_format({'bold': True})
                            worksheet.write_row('A1', nl, bold)

                            num_cols = len(nl)
                            rowS = 1
                            colS = 0
                            
                            nx = 0
                            data1 = []
                            data2 = []
                            data3 = []
                            data4 = []
                            data5 = []
                            data = []
                            for index, item in enumerate(lines):
                                if nx == 0:
                                    data1.append(item)
                                    nx += 1
                                    worksheet.write_column('A2', data1)
                                    if num_cols == 1:
                                        nx = 0
                                elif nx == 1:
                                    data2.append(item)
                                    nx += 1
                                    worksheet.write_column('B2', data2)
                                    if num_cols == 2:
                                        nx = 0
                                elif nx == 2:
                                    data3.append(item)
                                    nx += 1
                                    worksheet.write_column('C2', data3)
                                    if num_cols == 3:
                                        nx = 0
                                elif nx == 3:
                                    data4.append(item)
                                    nx += 1
                                    worksheet.write_column('D2', data4)
                                    if num_cols == 4:
                                        nx = 0
                                elif nx == 4:
                                    data5.append(item)
                                    nx = 0
                                    worksheet.write_column('E2', data5)
                                                                       
                            workbook.close()
                            finished = showinfo("Complete", "Your file has been sucessfully formatted and save as an Excel Spreadsheet.")
                            top.destroy()
                            
                        b2 = Button(top3, text="Ok", command=makeSheet)
                        b2.pack()
                b = Button(top2, text="Ok", command=storecols)
                b.pack()

            answer = askyesno("Format", "Would you like to format this file?", parent=top)
            if answer == True:
                formatFile()
            else:
                top.destroy()
            
    except FileNotFoundError:
        print("File was not selected, or not found.")
        top.destroy()

lbl1 = Label(window, text="This program can be used to format JSON Line files into text files.", font=("Helvetica", 14), pady=15)
lbl1.grid(row=0, column=0)
lbl2 = Label(window, text="JSON Line files are denoted by the .jl file extension.", font=("Helvetica", 12), pady=20)
lbl2.grid(row=1, column=0)
lbl3 = Label(window, text="To get started click the 'Open File' button on the bottom and locate the file.", font=("Helvetica", 12), pady=20)
lbl3.grid(row=2, column=0)
lbl4 = Label(window, text="After you have parsed a file you can format it to .xls or by pressing\n the 'Format' button.", font=("Helvetica", 12), pady=20)
lbl4.grid(row=3, column=0)

openButton = Button(window, text="Open File")
openButton['command'] = openFile
openButton.grid(row=4, column=0, sticky=SW, padx=50, pady=75)

formatButton = Button(window, text="Format")
formatButton['command'] = selectFileToFormat
formatButton.grid(row=4, column=0, sticky=S, padx=50, pady=75)

cancelButton = Button(window, text="Cancel", command=window.destroy)
cancelButton.grid(row=4, column=0, sticky=SE, padx=50, pady=75)

window.title("JSON Line Formatter")
window.geometry('560x400')
window.maxsize(560, 400)
window.minsize(560, 400)
window.mainloop()
