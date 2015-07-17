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
This scribt will copy the selected objects to all subsequent pages currently
existing in the document. No pages will be added yet.

FUTURE FEATURES
* The user should be presented with a dialog in order to control the script
    - Select a CSV file that holds the data that should be merge into 
      the document
    - Select how many rows should be merged into the document
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

class CONST:
    APP_NAME = 'Scribus Data Merger'

class DataMerger:
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

        numPages = scribus.pageCount()
        startingPage = scribus.currentPage()
        currentPage = startingPage + 1
        while (currentPage <= numPages):
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

def main(argv):
    merger = DataMerger()
    merger.run()    

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
