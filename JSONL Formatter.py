import tkinter as tk
import re, xlsxwriter
from tkinter import *
from tkinter import Tk
from tkinter.filedialog import askopenfilename, asksaveasfilename
from tkinter.messagebox import askyesno, showerror, showinfo

window = Tk()

def openFile():
    try:
        ftypes = [('JL file',"*.jl")]
        ttl = "Title"
        dir1 = 'T:\\'
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
                    fileText = match
                    display.insert(END, fileText + '\n')
                    #print(fileText)


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
        dir1 = 'T:\\'
        window.fileName = askopenfilename(filetypes = ftypes, initialdir = dir1, title = ttl)
        print(window.fileName)
        selectedFile = window.fileName
        top = Toplevel(window)
        with open(selectedFile, "r") as f:
            fileText = f.read()
            lines = fileText.splitlines()
            #print(lines)
            evli = []
            odli = []
            for index, line in enumerate(lines):
                if index % 2 == 0:
                    evli.append(line)
                else:
                    odli.append(line)
            #print(evli)
            #print(odli)
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
                            for name in name_list:
                                print(name.get())
                                name = name.get()
                            top3.destroy()
                            workbook = xlsxwriter.Workbook('%s.xlsx' % (selectedFile))
                            worksheet = workbook.add_worksheet()
                            if numcols == 1:
                                worksheet.write('A1', name)
                            elif numcols == 2:
                                worksheet.write('A1', name)
                                worksheet.write('B1', name)
                            elif numcols == 3:
                                worksheet.write('A1', name)
                                worksheet.write('B1', name)
                                worksheet.write('C1', name)
                            elif numcols == 4:
                                worksheet.write('A1', name)
                                worksheet.write('B1', name)
                                worksheet.write('C1', name)
                                worksheet.write('D1', name)
                            elif numcols == 5:
                                worksheet.write('A1', name)
                                worksheet.write('B1', name)
                                worksheet.write('C1', name)
                                worksheet.write('D1', name)
                                worksheet.write('E1', name)
                                
                            rowS = 1
                            colS = 0

                            for el in evli:
                                worksheet.write(rowS, colS, el)
                                rowS += 1
                            rowS = 1
                            colS = 0
                            for ol in odli:
                                worksheet.write(rowS, colS + 1, ol)
                                rowS += 1
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
