# Scribus-Data-Merger
Data merge tool that will help you populate a scribus file with data from a CSV file.

## Status
Currently the tool will only populate existing pages in the same document with data from the CSV file. You can select how many pages will be populated in the dialog.

## Usage
Before you run the script you must have a document open and at least one object selected. 

In the dialog select a CSV file from the file selector, select how many pages/rows to merge and click the *Run* button. You can choose how many pages the selection will be copied to (this is also the number of rows that will be read). After clicking *Run* button the script will copy the selected elements to the chosen number of existing subsequent pages of the document and replace variables with data from the CSV file.

### Structure of the Input File
The input file selected in the dialog must be a CSV file with the first row containing the name of each column. These header names are used to identify what cell to use when replacing text in the document.

### Variables
In any text frame you can specify variables or placeholders like this: `%VAR_employee%`. When the script runs it will replace the variable with the *content* of the cell with header name 'employee'.

## Setup
There are two ways to get the script to work:

1. Copy the script file (ScribusDataMerger.py) to the default extension folder `SCRIBUS_HOME\share\scripts\`, where `SCRIBUS_HOME` is the Scribus installation directory. 
	- After this you can run the script from the menu *Scripter->Scribus Scripts*
	- You might need to restart scribus in order for Scribus Data Merger to show up in the submenu.
1. From the menu select *Scripter->Run script...* and browse to the file with the *Run script* dialog.