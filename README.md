# Scribus-Data-Merger
Data merge tool that will help you populate a [Scribus](http://www.scribus.net/) file with data from a CSV file.

## Status
This is the first working version of the script. It will load a CSV file, create a new page pr. row in the CSV file and replace variables in text frames with data from the CSV file.

*Note that this only works with Scribus version 1.5.0.* If you need this functionality in Scribus version 1.4.X you should take a look at [ScribusGenerator](https://github.com/berteh/ScribusGenerator). ScribusGenerator offers slightly different functionality.

## Usage
Before you run the script you must have a document open and at least one (text frame) object selected. 

In the dialog select a CSV file from the file selector, select how many pages/rows to merge and click the *Run* button. The script will then create a new page for each row read from the CSV file and copy the selected elements to those pages and replace variables with data from the CSV file.

### Structure of the Input File
The input file selected in the dialog must be a CSV file with the first row containing the name of each column. These header names are used to identify what cell to use when replacing text in the document.

### Variables
In any text frame you can specify variables (placeholders) like this: `%VAR_employee%`. When the script runs it will replace the variable with the *content* of the cell with header name 'employee'.

## Setup
There are two ways to get the script to work:

1. Copy the script file (ScribusDataMerger.py) to the default extension folder `SCRIBUS_HOME\share\scripts\`, where `SCRIBUS_HOME` is the Scribus installation directory. 
	- After this you can run the script from the menu *Scripter->Scribus Scripts*
	- You might need to restart scribus in order for Scribus Data Merger to show up in the submenu.
1. From the menu select *Scripter->Run script...* and browse to the file with the *Run script* dialog.