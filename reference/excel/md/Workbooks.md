# Workbooks

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000208DB-0000-0000-C000-000000000046}  

A collection of all the Workbook objects that are currently open in the Microsoft Excel application.

**Remarks:** For more information about using a single Workbook object, see the Workbook object.

**Example:**

```vba
Workbooks.Close
```

## Properties (7)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Count As Long  (read-only)`  
  Returns a Long value that represents the number of objects in the collection.
- `Item As Workbook  (read-only)`  
  Returns a single object from a collection.
- `_NewEnum As IUnknown  (read-only)`
- `_Default As Workbook  (read-only)`

## Methods (8)

- `Add([Template As Variant]) As Workbook`  
  Creates a new workbook. The new workbook becomes the active workbook.
    - `Template As Variant` (optional): Determines how the new workbook is created. If this argument is a string specifying the name of an existing Microsoft Excel file, the new workbook is created with the specified file as a template. If this argument is a constant, the new workbook contains a single sheet of the specified type. Can be one of the following XlWBATemplate constants: xlWBATChart, xlWBATExcel4IntlMacroSheet, xlWBATExcel4MacroSheet, or xlWBATWorksheet. If this argument is omitted, Microsoft Excel creates a new workbook with a number of blank sheets (the number of sheets is set by the SheetsInNewWorkbook property).
- `Close()`  
  Closes the object.
- `Open(Filename As String, [UpdateLinks As Variant], [ReadOnly As Variant], [Format As Variant], [Password As Variant], [WriteResPassword As Variant], [IgnoreReadOnlyRecommended As Variant], [Origin As Variant], [Delimiter As Variant], [Editable As Variant], [Notify As Variant], [Converter As Variant], [AddToMru As Variant], [Local As Variant], [CorruptLoad As Variant]) As Workbook`  
  Opens a workbook.
    - `Filename As String` (required): String. The file name of the workbook to be opened.
    - `UpdateLinks As Variant` (optional): Specifies the way external references (links) in the file, such as the reference to a range in the Budget.xls workbook in the following formula =SUM([Budget.xls]Annual!C10:C25), are updated. If this argument is omitted, the user is prompted to specify how links will be updated. For more information about the values used by this parameter, see the Remarks section. If Microsoft Excel is opening a file in the WKS, WK1, or WK3 format and the UpdateLinks argument is 0, no charts are created; otherwise, Microsoft Excel generates charts from the graphs attached to the file.
    - `ReadOnly As Variant` (optional): True to open the workbook in read-only mode.
    - `Format As Variant` (optional): If Microsoft Excel opens a text file, this argument specifies the delimiter character. If this argument is omitted, the current delimiter is used. For more information about the values used by this parameter, see the Remarks section.
    - `Password As Variant` (optional): A string that contains the password required to open a protected workbook. If this argument is omitted and the workbook requires a password, the user is prompted for the password.
    - `WriteResPassword As Variant` (optional): A string that contains the password required to write to a write-reserved workbook. If this argument is omitted and the workbook requires a password, the user will be prompted for the password.
    - `IgnoreReadOnlyRecommended As Variant` (optional): True to have Microsoft Excel not display the read-only recommended message (if the workbook was saved with the Read-Only Recommended option).
    - `Origin As Variant` (optional): If the file is a text file, this argument indicates where it originated, so that code pages and Carriage Return/Line Feed (CR/LF) can be mapped correctly. Can be one of the following XlPlatform constants: xlMacintosh, xlWindows, or xlMSDOS. If this argument is omitted, the current operating system is used.
    - `Delimiter As Variant` (optional): If the file is a text file and the _Format_ argument is 6, this argument is a string that specifies the character to be used as the delimiter. For example, use Chr(9) for tabs, use "," for commas, use ";" for semicolons, or use a custom character. Only the first character of the string is used.
    - `Editable As Variant` (optional): If the file is a Microsoft Excel 4.0 add-in, this argument is True to open the add-in so that it is a visible window. If this argument is False or omitted, the add-in is opened as hidden, and it cannot be unhidden. This option does not apply to add-ins created in Microsoft Excel 5.0 or later. If the file is an Excel template, True to open the specified template for editing. False to open a new workbook based on the specified template. The default value is False.
    - `Notify As Variant` (optional): If the file cannot be opened in read/write mode, this argument is True to add the file to the file notification list. Microsoft Excel will open the file as read-only, poll the file notification list, and then notify the user when the file becomes available. If this argument is False or omitted, no notification is requested, and any attempts to open an unavailable file will fail.
    - `Converter As Variant` (optional): The index of the first file converter to try when opening the file. The specified file converter is tried first; if this converter does not recognize the file, all other converters are tried. The converter index consists of the row numbers of the converters returned by the FileConverters property.
    - `AddToMru As Variant` (optional): True to add this workbook to the list of recently used files. The default value is False.
    - `Local As Variant` (optional): True saves files against the language of Microsoft Excel (including control panel settings). False (default) saves files against the language of Visual Basic for Applications (VBA) (which is typically United States English unless the VBA project where Workbooks.Open is run from is an old internationalized XL5/95 VBA project).
    - `CorruptLoad As Variant` (optional): Can be one of the following constants: xlNormalLoad, xlRepairFile and xlExtractData. The default behavior if no value is specified is xlNormalLoad, and does not attempt recovery when initiated through the OM.
- `OpenText(Filename As String, [Origin As Variant], [StartRow As Variant], [DataType As Variant], [TextQualifier As XlTextQualifier], [ConsecutiveDelimiter As Variant], [Tab As Variant], [Semicolon As Variant], [Comma As Variant], [Space As Variant], [Other As Variant], [OtherChar As Variant], [FieldInfo As Variant], [TextVisualLayout As Variant], [DecimalSeparator As Variant], [ThousandsSeparator As Variant], [TrailingMinusNumbers As Variant], [Local As Variant])`  
  Loads and parses a text file as a new workbook with a single sheet that contains the parsed text-file data.
    - `Filename As String` (required): Specifies the file name of the text file to be opened and parsed.
    - `Origin As Variant` (optional): Specifies the origin of the text file. Can be one of the following XlPlatform constants: xlMacintosh, xlWindows, or xlMSDOS. Additionally, this could be an integer representing the code page number of the desired code page. For example, "1256" would specify that the encoding of the source text file is Arabic (Windows). If this argument is omitted, the method uses the current setting of the File Origin option in the Text Import Wizard.
    - `StartRow As Variant` (optional): The row number at which to start parsing text. The default value is 1.
    - `DataType As Variant` (optional): Specifies the column format of the data in the file. Can be one of the following XlTextParsingType constants: xlDelimited or xlFixedWidth. If this argument is not specified, Microsoft Excel attempts to determine the column format when it opens the file.
    - `TextQualifier As XlTextQualifier` (optional): Specifies the text qualifier.
    - `ConsecutiveDelimiter As Variant` (optional): True to have consecutive delimiters considered one delimiter. The default is False.
    - `Tab As Variant` (optional): True to have the tab character be the delimiter (DataType must be xlDelimited). The default value is False.
    - `Semicolon As Variant` (optional): True to have the semicolon character be the delimiter (DataType must be xlDelimited). The default value is False.
    - `Comma As Variant` (optional): True to have the comma character be the delimiter (DataType must be xlDelimited). The default value is False.
    - `Space As Variant` (optional): True to have the space character be the delimiter (DataType must be xlDelimited). The default value is False.
    - `Other As Variant` (optional): True to have the character specified by the _OtherChar_ argument be the delimiter (DataType must be xlDelimited). The default value is False.
    - `OtherChar As Variant` (optional): Required if _Other_ is True. Specifies the delimiter character when _Other_ is True. If more than one character is specified, only the first character of the string is used; the remaining characters are ignored.
    - `FieldInfo As Variant` (optional): An array containing parse information for individual columns of data. The interpretation depends on the value of _DataType_. When the data is delimited, this argument is an array of two-element arrays, with each two-element array specifying the conversion options for a particular column. The first element is the column number (1-based), and the second element is one of the XlColumnDataType constants specifying how the column is parsed.
    - `TextVisualLayout As Variant` (optional): The visual layout of the text.
    - `DecimalSeparator As Variant` (optional): The decimal separator that Microsoft Excel uses when recognizing numbers. The default setting is the system setting.
    - `ThousandsSeparator As Variant` (optional): The thousands separator that Excel uses when recognizing numbers. The default setting is the system setting.
    - `TrailingMinusNumbers As Variant` (optional): Specify True if numbers with a minus character at the end should be treated as negative numbers. If False or omitted, numbers with a minus character at the end are treated as text.
    - `Local As Variant` (optional): Specify True if regional settings of the machine should be used for separators, numbers and data formatting.
- `OpenDatabase(Filename As String, [CommandText As Variant], [CommandType As Variant], [BackgroundQuery As Variant], [ImportDataAs As Variant]) As Workbook`  
  Returns a Workbook object representing a database.
    - `Filename As String` (required): The connection string that contains the location and file name of the database.
    - `CommandText As Variant` (optional): The command text of the query.
    - `CommandType As Variant` (optional): The command type of the query. Specify one of the constants of the XlCmdType enumeration: xlCmdCube, xlCmdList, xlCmdSql, xlCmdTable, and xlCmdDefault.
    - `BackgroundQuery As Variant` (optional): This parameter is a variant data type but you can only pass a Boolean value. If you pass True, the query is performed in the background (asynchronously). The default value is False.
    - `ImportDataAs As Variant` (optional): This parameter uses one of the values of the XlImportDataAs enumeration. The two values of this enum are xlPivotTableReport and xlQueryTable. Pass one of these values to return the data as a PivotTable or QueryTable. The default value is xlQueryTable.
- `CheckOut(Filename As String)`  
  Returns a String representing a specified workbook from a server to a local computer for editing.
    - `Filename As String` (required): The name of the file to check out.
- `CanCheckOut(Filename As String) As Boolean`  
  True if Microsoft Excel can check out a specified workbook from a server. Read/write Boolean.
    - `Filename As String` (required): The name of the file to check out.
- `OpenXML(Filename As String, [Stylesheets As Variant], [LoadOption As Variant]) As Workbook`  
  Opens an XML data file. Returns a Workbook object.
    - `Filename As String` (required): The name of the file to open.
    - `Stylesheets As Variant` (optional): Either a single value or an array of values that specify which XSL Transformation (XSLT) stylesheet processing instructions to apply.
    - `LoadOption As Variant` (optional): Specifies how Excel opens the XML data file. Can be one of the XlXmlLoadOption constants.
