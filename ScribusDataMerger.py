#!/usr/bin/env python
# -*- coding: utf-8 -*-

# File: ScribusDataMerge.py
# © 2015.07.16 Bjarne Svanberg
# This version 2015.07.16

# The MIT License (MIT)

# Copyright (c) 2015 Bjarne Svanberg

# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:

# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.

# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.


"""
USAGE
You must have a document open. Select a number of objects. Run the script. 
After selecting a file in the dialog and clicking the Run button the scribt 
will copy the selected objects to selected number of subsequent pages 
currently existing in the document. No pages will be added yet.

FUTURE FEATURES
* The user should be presented with a dialog in order to control the script
    - Select starting row
* The script should create a new page per row of data in the CSV input file
* The script should replace any variables %VAR_name% with the content of 
  a cell with the header 'name'

"""

import sys
 
try:
    # Please do not use 'from scribus import *' . If you must use a 'from import',
    # Do so _after_ the 'import scribus' and only import the names you need, such
    # as commonly used constants.
    import scribus
except ImportError,err:
    print "This Python script is written for the Scribus scripting interface."
    print "It can only be run from within Scribus."
    sys.exit(1)

import Tkinter
from Tkinter import Frame, LabelFrame, Label, Entry, Button, StringVar, OptionMenu, Checkbutton, IntVar
import tkFileDialog
import tkMessageBox

class CONST:
    APP_NAME = 'Scribus Data Merger'
    EMPTY = ''

class DataMerger:
    def __init__(self, dataObject):
        self.__dataObject = dataObject

    def run(self):
        selCount = scribus.selectionCount()
        if selCount == 0:
            scribus.messageBox('Scribus Data Merger- Usage Error',
                               "There is no objects selected.\nPlease try again.",
                               scribus.ICON_WARNING, scribus.BUTTON_OK)
            sys.exit(2)
        

        # Create a list with the names of the selected objects:
        selectedObjects = []

        # Loop through the selected objects and put their names into the list selectedObjects
        o = 0 
        while (o < selCount):
            selectedObjects.append(scribus.getSelectedObject(o))
            o = o + 1

        startingPage = scribus.currentPage()
        numPages = scribus.pageCount()
        lastPage = numPages # Default value
        if(self.__dataObject.getNumberOfLinesToMerge() != 'All'):
            lastPage = int(self.__dataObject.getNumberOfLinesToMerge()) + startingPage
        lastPage = min(lastPage, numPages) # This will prevent the script from trying to paste objects to non-existing pages
        currentPage = startingPage + 1
        while (currentPage <= lastPage):
            for selectedObject in selectedObjects:
                scribus.gotoPage(startingPage) # Set the working page to the we want to copy objects from 
                scribus.copyObject(selectedObject)

                # self.info('Object', selectedObject)
                scribus.gotoPage(currentPage)
                scribus.pasteObject()
                # self.info('newobject: ', newobject)

            scribus.docChanged(1)
            newPageObejcts = scribus.getAllObjects()


            for pastedObject in newPageObejcts: # Loop through all the items on the current page
                objType = scribus.getObjectType(pastedObject)
                # self.info('pastedObject', pastedObject)
                # self.info('objType', objType)
                text = 'unchanged'
                if(objType == 'TextFrame'):
                    # text = scribus.getText(pastedObject) 
                    text = scribus.getAllText(pastedObject) # This should have used getText but getText does not return the text values of just pasted objects
                    # Todo: Insert text from the CSV file in the text variable before it is set on the pastedObject
                    scribus.setText(text, pastedObject)
                if(objType == 'ImageFrame'):
                    text = scribus.getImageFile(pastedObject)
                    # Todo: Find out if it is possible to replace text in the ImageFile property
                          
            currentPage = currentPage + 1
     
        scribus.setRedraw(1)
        scribus.docChanged(1)
        scribus.messageBox("Merge Completed", "Merge Completed", icon=scribus.ICON_INFORMATION, button1=scribus.BUTTON_OK)


    def info(self, text, var):
        """ Shorthand method for showing information in a dialog box """
        scribus.messageBox("Info", text + ": " + str(var), icon=scribus.ICON_INFORMATION, button1=scribus.BUTTON_OK)

class MergerDataObject:
    # Data Object for transfering the settings made by the user on the UI
    def __init__(self):
        self.__dataSourceFile = CONST.EMPTY
        self.__numberOfLinesToMerge = '0'
    
    # Get
    def getDataSourceFile(self):
        return self.__dataSourceFile

    def getNumberOfLinesToMerge(self):
        return self.__numberOfLinesToMerge
    
    # Set
    def setDataSourceFile(self, fileName):
        self.__dataSourceFile = fileName
    
    def setNumberOfLinesToMerge(self, stringValue):
        self.__numberOfLinesToMerge = stringValue

class MergerController:
    # Controler being the bridge between UI and Logic.
    def __init__(self, root):
        self.__dataSourceFileEntryVariable = StringVar()
        self.__scribusSourceFileEntryVariable = StringVar()
        self.__outputDirectoryEntryVariable = StringVar()
        self.__outputFileNameEntryVariable = StringVar()
        self.__linesToMerge = ["All", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20"]
        self.__selectedLinesToMerge = StringVar()
        self.__selectedLinesToMerge.set('All') # The default value for the dropdown box
        self.__root = root

    def getDataSourceFileEntryVariable(self):
        return self.__dataSourceFileEntryVariable

    def dataSourceFileEntryVariableHandler(self):
        result = tkFileDialog.askopenfilename(title='Choose...', defaultextension='.csv', filetypes=[('CSV', '*.csv *.CSV')])
        if result:
            self.__dataSourceFileEntryVariable.set(result)

    def createDataObject(self):
        # Collect the settings the user has made and build the Data Object
        result = MergerDataObject()
        result.setDataSourceFile(self.__dataSourceFileEntryVariable.get())
        result.setNumberOfLinesToMerge(self.__selectedLinesToMerge.get())
        return result

    def getSelectedNumberOfLines(self):
        return self.__selectedLinesToMerge

    def getNumerOfLinesList(self):
        return self.__linesToMerge

    def quit(self):
        self.__root.destroy()

    def buttonCancelHandler(self):
        self.quit()

    def buttonOkHandler(self):
        if (self.__dataSourceFileEntryVariable.get() != CONST.EMPTY): 
            dataObject = self.createDataObject()
            merger = DataMerger(dataObject)
            merger.run()
            self.quit()
        else:
            tkMessageBox.showerror(title='Validation failed', message='Please check if all settings have been set correctly!')

class MergerDialog:
    # The UI to configure the settings by the user
    def __init__(self, root, ctrl):
        self.__root = root
        self.__ctrl = ctrl
    
    def show(self):
        self.__root.title(CONST.APP_NAME)
        mainFrame = Frame(self.__root)
        # Configure main frame and make Dialog stretchable (to EAST and WEST)
        top = mainFrame.winfo_toplevel()
        top.rowconfigure(0, weight=1)
        top.columnconfigure(0, weight=1)
        mainFrame.rowconfigure(0, weight=1)
        mainFrame.columnconfigure(0, weight=1)
        mainFrame.grid(sticky='ew')

        # Configure the frames to hold the controls
        # - Frame with input settings
        inputFrame = LabelFrame(mainFrame, text='Input Settings')
        inputFrame.columnconfigure(1, weight=1)
        inputFrame.grid(column=0, row=0, padx=5, pady=5, sticky='ew')
        # - Frame with output settings
        outputFrame = LabelFrame(mainFrame, text='Output Settings')
        outputFrame.columnconfigure(1, weight=1)
        outputFrame.grid(column=0, row=1, padx=5, pady=5, sticky='ew')
        # - Frame with buttons at the bottom of the dialog
        buttonFrame = Frame(mainFrame)
        buttonFrame.grid(column=0, row=2, padx=5, pady=5, sticky='e')

        # Controls for input
        dataSourceFileLabel = Label(inputFrame, text='Data File:', width=15, anchor='w')
        dataSourceFileLabel.grid(column=0, row=1, padx=5, pady=5, sticky='w')
        dataSourceFileEntry = Entry(inputFrame, width=70, textvariable=self.__ctrl.getDataSourceFileEntryVariable())
        dataSourceFileEntry.grid(column=1, row=1, padx=5, pady=5, sticky='ew')
        dataSourceFileButton = Button(inputFrame, text='...', command=self.__ctrl.dataSourceFileEntryVariableHandler)
        dataSourceFileButton.grid(column=2, row=1, padx=5, pady=5, sticky='e')

        # Controls for output
        numberOfDataLinesToMergeLabel = Label(outputFrame, text='Number of data lines to merge into this document:', width=15, anchor='w')
        numberOfDataLinesToMergeLabel.grid(column=0, row=2, padx=5, pady=5, sticky='w')
        # numberOfDataLinesToMergeListBox = OptionMenu(outputFrame, self.__ctrl.getSelectedNumberOfLines(), tuple(self.__ctrl.getNumerOfLinesList())) # TODO: implement these two functions in the controller
        numberOfDataLinesToMergeListBox = apply(OptionMenu, (outputFrame, self.__ctrl.getSelectedNumberOfLines()) + tuple(self.__ctrl.getNumerOfLinesList()))
        numberOfDataLinesToMergeListBox.grid(column=1, row=2, padx=5, pady=5, sticky='w')

        # Buttons
        cancelButton = Button(buttonFrame, text='Cancel', width=10, command=self.__ctrl.buttonCancelHandler)
        cancelButton.grid(column=1, row=0, padx=5, pady=5, sticky='e')
        runButton = Button(buttonFrame, text='Run', width=10, command=self.__ctrl.buttonOkHandler)
        runButton.grid(column=2, row=0, padx=5, pady=5, sticky='e')

        # Show the dialog
        mainFrame.grid()
        self.__root.grid()
        self.__root.mainloop()

def main(argv):
    root = Tkinter.Tk()
    ctrl = MergerController(root)
    dlg = MergerDialog(root, ctrl)
    dlg.show()

def main_wrapper(argv):
    try:
        if(scribus.haveDoc()):
            scribus.setRedraw(False)
            scribus.statusMessage(CONST.APP_NAME)
            scribus.progressReset()
            main(argv)
        else:
            scribus.messageBox('Usage Error', 'You need a Document open', icon=scribus.ICON_WARNING, button1=scribus.BUTTON_OK)
            sys.exit(2)
    finally:
        # Exit neatly even if the script terminated with an exception,
        # so we leave the progress bar and status bar blank and make sure
        # drawing is enabled.
        if scribus.haveDoc():
            scribus.setRedraw(True)
        scribus.statusMessage("")
        scribus.progressReset()
 
# This code detects if the script is being run as a script, or imported as a module.
# It only runs main() if being run as a script. This permits you to import your script
# and control it manually for debugging.
if __name__ == '__main__':
    main_wrapper(sys.argv)
