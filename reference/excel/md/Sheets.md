# Sheets

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000208D7-0000-0000-C000-000000000046}  

A collection of all the sheets in the specified or active workbook.

**Remarks:** The Sheets collection can contain Chart or Worksheet objects. The Sheets collection is useful when you want to return sheets of any type. If you need to work with sheets of only one type, see the object topic for that sheet type.

**Example:**

```vba
Sheets.PrintOut
```

## Properties (10)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Count As Long  (read-only)`  
  Returns a Long value that represents the number of objects in the collection.
- `Item As Object  (read-only)`  
  Returns a single object from a collection.
- `_NewEnum As IUnknown  (read-only)`
- `HPageBreaks As HPageBreaks  (read-only)`  
  Returns an HPageBreaks collection that represents the horizontal page breaks on the sheet. Read-only.
- `VPageBreaks As VPageBreaks  (read-only)`  
  Returns a VPageBreaks collection that represents the vertical page breaks on the sheet. Read-only.
- `Visible As Variant  (read/write)`  
  Returns or sets a Variant value that determines whether the object is visible.
- `_Default As Object  (read-only)`

## Methods (9)

- `Add([Before As Variant], [After As Variant], [Count As Variant], [Type As Variant]) As Object`  
  Creates a new worksheet, chart, or macro sheet. The new worksheet becomes the active sheet.
    - `Before As Variant` (optional): An object that specifies the sheet before which the new sheet is added.
    - `After As Variant` (optional): An object that specifies the sheet after which the new sheet is added.
    - `Count As Variant` (optional): The number of sheets to be added. The default value is the number of selected sheets.
    - `Type As Variant` (optional): Specifies the sheet type. Can be one of the following XlSheetType constants: xlWorksheet, xlChart, xlExcel4MacroSheet, or xlExcel4IntlMacroSheet. If you are inserting a sheet based on an existing template, specify the path to the template. The default value is xlWorksheet.
- `Copy([Before As Variant], [After As Variant])`  
  Copies the sheet to another location in the workbook.
    - `Before As Variant` (optional): The sheet before which the copied sheet will be placed. You cannot specify _Before_ if you specify _After_.
    - `After As Variant` (optional): The sheet after which the copied sheet will be placed. You cannot specify _After_ if you specify _Before_.
- `Delete()`  
  Deletes the object.
- `FillAcrossSheets(Range As Range, [Type As XlFillWith])`  
  Copies a range to the same area on all other worksheets in a collection.
    - `Range As Range` (required): The range to fill on all the worksheets in the collection. The range must be from a worksheet within the collection.
    - `Type As XlFillWith` (optional): Specifies how to copy the range.
- `Move([Before As Variant], [After As Variant])`  
  Moves the sheet to another location in the workbook.
    - `Before As Variant` (optional): The sheet before which the moved sheet will be placed. You cannot specify _Before_ if you specify _After_.
    - `After As Variant` (optional): The sheet after which the moved sheet will be placed. You cannot specify _After_ if you specify _Before_.
- `PrintPreview([EnableChanges As Variant])`  
  Shows a preview of the object as it would look when printed.
    - `EnableChanges As Variant` (optional): Pass a Boolean value to specify if the user can change the margins and other page setup options available in print preview.
- `Select([Replace As Variant])`  
  Selects the object.
    - `Replace As Variant` (optional): Used only with sheets. True to replace the current selection with the specified object. False to extend the current selection to include any previously selected objects and the specified object.
- `PrintOut([From As Variant], [To As Variant], [Copies As Variant], [Preview As Variant], [ActivePrinter As Variant], [PrintToFile As Variant], [Collate As Variant], [PrToFileName As Variant], [IgnorePrintAreas As Variant])`  
  Prints the object.
    - `From As Variant` (optional): The number of the page at which to start printing. If this argument is omitted, printing starts at the beginning.
    - `To As Variant` (optional): The number of the last page to print. If this argument is omitted, printing ends with the last page.
    - `Copies As Variant` (optional): The number of copies to print. If this argument is omitted, one copy is printed.
    - `Preview As Variant` (optional): True to have Microsoft Excel invoke print preview before printing the object. False (or omitted) to print the object immediately.
    - `ActivePrinter As Variant` (optional): Sets the name of the active printer.
    - `PrintToFile As Variant` (optional): True to print to a file. If _PrToFileName_ is not specified, Excel prompts the user to enter the name of the output file.
    - `Collate As Variant` (optional): True to collate multiple copies.
    - `PrToFileName As Variant` (optional): If _PrintToFile_ is set to True, this argument specifies the name of the file that you want to print to.
    - `IgnorePrintAreas As Variant` (optional): True to ignore print areas and print the entire object.
- `Add2([Before As Variant], [After As Variant], [Count As Variant], [NewLayout As Variant]) As Object`  
  This method is only implemented for the Charts collection object and will produce a run-time error if used on the Sheets and Worksheets objects.
    - `Before As Variant` (optional): An object that specifies the sheet before which the new sheet is added.
    - `After As Variant` (optional): An object that specifies the sheet after which the new sheet is added.
    - `Count As Variant` (optional): The number of sheets to be added. The default value is one.
    - `NewLayout As Variant` (optional): The layout of the new worksheet.
