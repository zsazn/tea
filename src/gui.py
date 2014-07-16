#! /usr/bin/env python
#-*- encoding = UTF-8 -*-
__author__ = 'Mark Z. Zhou'
__version__ = '1.0'

import core
from Tkinter import * 

namepass = ''
class Application(Frame):
    entries = []
    buttons = []
    labelList = ['1. Choose Faculty', '2. Choose semester', '3. Choose Course', 'or choose CRN', 'Faculty Name List', 'Semester List', 'Course List', 'CRN List']         

    def createWidgets(self):
        # for item in self.labelList:
            # row = Frame(self)
            # row.pack(side = TOP, fill = X)
            # Label(row, text = item, width = 15, height = 2).pack(side = LEFT)
            # entry = Entry(row, bg = 'white')
            # entry.pack(side = RIGHT, expand = YES, fill = X)
            # entry.bind('<PRINTFORM>', (lambda event:self.getValue()))
            # self.entries.append(entry)
        for i in xrange(4):
            row = Frame(self)
            row.pack(side = TOP, fill = X)
            Label(row, text = self.labelList[i], width = 15, height = 2).pack(side = LEFT)
            button = Button(row, text = self.labelList[i+4], width = 15, height = 1)
            button.pack(side = RIGHT)
            self.buttons.append(button)
            
        
        self.PRINTFORM = Button(self,
                                text = 'PRINT',
                                command = self.printForm,
                                width = 10,
                                bd = 4,
                                relief = RIDGE).pack(side = LEFT)
        self.RESET = Button(self,
                            text = 'RESET',
                            command = self.clearData,
                            width = 10,
                            bd = 4,
                            relief = RIDGE).pack(side = LEFT)
        self.QUIT = Button(self,
                           text = 'QUIT',
                           command = self.quit,
                           width = 10,
                           bd = 4,
                           relief = RIDGE).pack(side = LEFT)

#    def getListbox(self, c):
        

    def getValue(self):
        result = []
        for entry in self.entries:
            result.append(entry.get())
        namepass = result[0]
        print namepass
        
    def printForm(self):
        self.getValue()

    def clearData(self):
        for entry in self.entries:
            entry.delete(0,END)
        
    def __init__(self, master = None):
        Frame.__init__(self, master)
        self.pack(side=TOP)
        self.createWidgets()
def main():
    root = Tk()
    root.title('Teaching Evaluation Report Generator')
    root.geometry('350x150')
    app = Application(master = root)
    app.mainloop()
    root.destroy()

if __name__ == '__main__':
    main()
