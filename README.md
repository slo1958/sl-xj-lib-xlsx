# sl-xj-xlsx


The library provides support to read xlsx files (and write xlsx files later)

About Xojo version: tested with Xojo 2024 release 4.1 on Mac.

## Read XLSX files

The library contains a set of classes to read a xlsx file. 

The main purpose is to get values from cells. Most style related information are ignored, excepted when those information are required to properly guess the type of the value stored in cell. 

Two test XLSX files are included. Those files were created with WPS Office or OnlyOffice. 

An example window shows:

- how style/format are used to guess the content of a few cells
- a sheet used as a form from example workbook 1
- tabular sheets stored in example workbook 2

### Examples

This example loads a workbook and all its dependencies in memory, using the system provided temporary folder as work area:

```xojo
var myWorkbook as clWorkbook =  new clWorkbook(myXLSXFile)
var sheets() as string =  myWorkbook.GetSheetNames

// Obtain sheet data
var sheet as clWorksheet =  myWorkbook.GetSheet(SheetName)

```

This example load a workbook in memory, the worksheet data are loaded uploaded on request:

```xojo
var myWorkbook as clWorkbook =  new clWorkbook(myXLSXFile)
var sheets() as string =  myWorkbook.GetSheetNames

// Obtain sheet data, this will cause the library to load the worksheet
var sheet as clWorksheet =  myWorkbook.GetSheet(SheetName)

// Do something with the sheet

// Remove the sheet data from memory
// Remember the library needs to load the sheet again from XML if access to the sheet is requested

myWorkbook.DropSheetFromMemory(SheetName)

// nil the pointer to zero the refcount
Sheet = nil

```

Get the formatted value of a cell (as a string):

```xojo

Var tempUsingAddress as string = worksheet.GetCell("B12").GetValueAsString(workbook)

Var tempUsingRowCol as string = worksheet.GetCell(12,2).GetValueAsString(workbook)

```
Note that the method returning the string value requires access to the workbook.


Get the value as a date:

```xojo
// Get a date value

Var cell as clCell = worksheet.GetCell("B12”)

If cell.GuessType(workbook) = clCell.GuessedType.Date then
	myDateTime = cell.GetValueAsDateTime(Workbook)
end if

```

Get the value as a number:

```xojo
// Get a number value

Var cell as clCell = worksheet.GetCell("B12”)

If cell.GuessType(workbook) = clCell.GuessedType.Number then
	myNumber = cell.GetValueAsNumber(Workbook)
end if

// or 

If cell.GuessType(workbook) = clCell.GuessedType.General then
	myNumber = cell.GetValueAsNumber(Workbook)
end if

```

### UI related components

The library offers some UI related components

- a container control to display the content of a sheet, the sheet is selected by the user via a popup menu
- a method to populate a listbox from a sheet
- a method to populate a popup menu with the name of the sheet


### Supported fatures
Creates the standard default number format (for number, percentage, dates,…)


Properly recognise a date format (and assumes the value should be treated as a date offset) for non localised format strings



### Unsupported features and limitations
Language specific format strings are skipped. For example, the following format for 1973-11-27 will results with the same formatted string

```xojo
Format :  [$-409]mmmm dd yyyy
Expected formatted value : November 27 1973 
Displayed formatted value :  November 27 1973

format:  [$-804]mmmm dd yyyy
Expected formatted value : 十一月 27 1973
Displayed formatted value :  November 27 1973

```

Similarly, the characters 年 月 日 are not recognised as indicating a date format.

Day names and month names used when formatting dates are taken from` ` hard-coded English names.

Time formats are not handled.

There are no calculation engine.  

The clCell() stores the formula associated with the cell only if that formula is associated with the cell in the xml file. Formula are not propagated 

More test cases must be included.

## Write XLSX files
In dev


## Adding the library to your project

- open the project test-lib-xlsx
- locate the folder ‘lib-xlsx’
- copy the folder
- paste in your project
- Xojo may expand the folder after pasting, you can collapse it
- If you want to display the content of data table in a form, for example for debugging purposes, repeat this operation for the folder ‘lib-xlsx-ui-support’
