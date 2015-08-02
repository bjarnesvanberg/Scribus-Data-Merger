# Scribus-Data-Merger
Data merge tool that will help you populate a scribus file with data from a CSV file.

## Status
Currently the tool will only copy selected objects to subsequent pages in the same document. You can select how many pages it will be dublicated to in the dialog.

## Usage
Before you run the script you must have a document open and at least one object selected. 

In the dialog select a file from the file selector and click the *Run* button. *Currently the selected file does not have any impact on the behavior of the script but it is still required to fill out.* You can choose how many pages the selection will be copied to. After clicking *Run* button the script will copy the selected elements to the chosen number of existing subsequent pages of the document.

## Setup
There are two ways to get the script to work:

1. Copy the script file (ScribusDataMerger.py) to the default extension folder `SCRIBUS_HOME\share\scripts\`, where `SCRIBUS_HOME` is the Scribus installation directory. 
	- After this you can run the script from the menu *Scripter->Scribus Scripts*
	- You might need to restart scribus in order for Scribus Data Merger to show up in the submenu.
1. From the menu select *Scripter->Run script...* and browse to the file with the *Run script* dialog.