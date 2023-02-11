# JavaExcelHandler

A Java program that creates a backup copy of an Excel file, displays its metadata, and removes worksheet/workbook write/edit protection.

This project was built with: 
* JavaFX for the user interface.
* Apache POI 5.2.3 libraries for classes dealing with Excel file manipulation.

## Features 

* Creates a backup of an Excel file to the same directory.
* Displays document and extended properties of a selected Excel file in a new JavaFX window.
* Removes passwords set at the worksheet or workbook level.

## Limitations
This program is designed to handle .xlsx and .xls files only and may not work with other types of Excel files (e.g. xlsm). 
It also does not handle password protection set at the file level.

## Contributions
Feel free to contribute by creating pull requests or reporting bugs.