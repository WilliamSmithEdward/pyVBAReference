# Worksheet

**Type:** Class  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00020820-0000-0000-C000-000000000046}  

Represents a worksheet.

**Remarks:** The Worksheet object is a member of the Worksheets collection. The Worksheets collection contains all the Worksheet objects in a workbook. The Worksheet object is also a member of the Sheets collection. The Sheets collection contains all the sheets in the workbook (both chart sheets and worksheets).

**Example:**

```vba
Worksheets(1).Visible = False
```

## Properties (58)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `CodeName As String  (read-only)`  
  Returns the code name for the object. Read-only String.
- `_CodeName As String  (read/write)`
- `Index As Long  (read-only)`  
  Returns a Long value that represents the index number of the object within the collection of similar objects.
- `Name As String  (read/write)`  
  Returns or sets a String value that represents the object name.
- `Next As Object  (read-only)`  
  Returns a Worksheet object that represents the next sheet.
- `PageSetup As PageSetup  (read-only)`  
  Returns a PageSetup object that contains all the page setup settings for the specified object. Read-only.
- `Previous As Object  (read-only)`  
  Returns a Worksheet object that represents the previous sheet.
- `ProtectContents As Boolean  (read-only)`  
  True if the contents of the sheet are protected. This protects the individual cells. To turn on content protection, use the Protect method with the _Contents_ argument set to True. Read-only Boolean.
- `ProtectDrawingObjects As Boolean  (read-only)`  
  True if shapes are protected. To turn on shape protection, use the Protect method with the _DrawingObjects_ argument set to True. Read-only Boolean.
- `ProtectionMode As Boolean  (read-only)`  
  True if user-interface-only protection is turned on. To turn on user interface protection, use the Protect method with the _UserInterfaceOnly_ argument set to True. Read-only Boolean.
- `ProtectScenarios As Boolean  (read-only)`  
  True if the worksheet scenarios are protected. Read-only Boolean.
- `Visible As XlSheetVisibility  (read/write)`  
  Returns or sets an XlSheetVisibility value that determines whether the object is visible.
- `Shapes As Shapes  (read-only)`  
  Returns a Shapes collection that represents all the shapes on the worksheet. Read-only.
- `TransitionExpEval As Boolean  (read/write)`  
  True if Microsoft Excel uses Lotus 1-2-3 expression evaluation rules for the worksheet. Read/write Boolean.
- `AutoFilterMode As Boolean  (read/write)`  
  True if the AutoFilter drop-down arrows are currently displayed on the sheet. This property is independent of the FilterMode property. Read/write Boolean.
- `EnableCalculation As Boolean  (read/write)`  
  True if Microsoft Excel automatically recalculates the worksheet when necessary. False if Excel doesn't recalculate the sheet. Read/write Boolean.
- `Cells As Range  (read-only)`  
  Returns a Range object that represents all the cells on the worksheet (not just the cells that are currently in use).
- `CircularReference As Range  (read-only)`  
  Returns a Range object that represents the range containing the first circular reference on the sheet, or returns Nothing if there's no circular reference on the sheet. The circular reference must be removed before calculation can proceed.
- `Columns As Range  (read-only)`  
  Returns a Range object that represents all the columns on the specified worksheet.
- `ConsolidationFunction As XlConsolidationFunction  (read-only)`  
  Returns the function code used for the current consolidation. Can be one of the constants of XlConsolidationFunction. Read-only Long.
- `ConsolidationOptions As Variant  (read-only)`  
  Returns a three-element array of consolidation options, as shown in the following table. If the element is True, that option is set. Read-only Variant.
- `ConsolidationSources As Variant  (read-only)`  
  Returns an array of string values that name the source sheets for the worksheet's current consolidation. Returns Empty if there's no consolidation on the sheet. Read-only Variant.
- `EnableAutoFilter As Boolean  (read/write)`  
  True if AutoFilter arrows are enabled when user-interface-only protection is turned on. Read/write Boolean.
- `EnableSelection As XlEnableSelection  (read/write)`  
  Returns or sets what can be selected on the sheet. Read/write XlEnableSelection.
- `EnableOutlining As Boolean  (read/write)`  
  True if outlining symbols are enabled when user-interface-only protection is turned on. Read/write Boolean.
- `EnablePivotTable As Boolean  (read/write)`  
  True if PivotTable controls and actions are enabled when user-interface-only protection is turned on. Read/write Boolean.
- `FilterMode As Boolean  (read-only)`  
  True if the worksheet is in the filter mode. Read-only Boolean.
- `Names As Names  (read-only)`  
  Returns a Names collection that represents all the worksheet-specific names (names defined with the "WorksheetName!" prefix). Read-only Names object.
- `Outline As Outline  (read-only)`  
  Returns an Outline object that represents the outline for the specified worksheet. Read-only.
- `Range As Range  (read-only)`  
  Returns a Range object that represents a cell or a range of cells.
- `Rows As Range  (read-only)`  
  Returns a Range object that represents all the rows on the specified worksheet.
- `ScrollArea As String  (read/write)`  
  Returns or sets the range where scrolling is allowed, as an A1-style range reference. Cells outside the scroll area cannot be selected. Read/write String.
- `StandardHeight As Double  (read-only)`  
  Returns the standard (default) height of all the rows on the worksheet, in points. Read-only Double.
- `StandardWidth As Double  (read/write)`  
  Returns or sets the standard (default) width of all the columns on the worksheet. Read/write Double.
- `TransitionFormEntry As Boolean  (read/write)`  
  True if Microsoft Excel uses Lotus 1-2-3 formula entry rules for the worksheet. Read/write Boolean.
- `Type As XlSheetType  (read-only)`  
  Returns an XlSheetType value that represents the worksheet type.
- `UsedRange As Range  (read-only)`  
  Returns a Range object that represents the used range on the specified worksheet. Read-only.
- `HPageBreaks As HPageBreaks  (read-only)`  
  Returns an HPageBreaks collection that represents the horizontal page breaks on the sheet. Read-only.
- `VPageBreaks As VPageBreaks  (read-only)`  
  Returns a VPageBreaks collection that represents the vertical page breaks on the sheet. Read-only.
- `QueryTables As QueryTables  (read-only)`  
  Returns the QueryTables collection that represents all the query tables on the specified worksheet. Read-only.
- `DisplayPageBreaks As Boolean  (read/write)`  
  True if page breaks (both automatic and manual) on the specified worksheet are displayed. Read/write Boolean.
- `Comments As Comments  (read-only)`  
  Returns a Comments collection that represents all the comments for the specified worksheet. Read-only.
- `Hyperlinks As Hyperlinks  (read-only)`  
  Returns a Hyperlinks collection that represents the hyperlinks for the worksheet.
- `DisplayRightToLeft As Boolean  (read/write)`  
  True if the specified worksheet is displayed from right to left instead of from left to right. False if the object is displayed from left to right. Read-only Boolean.
- `Tab As Tab  (read-only)`  
  Returns a Tab object for a worksheet.
- `MailEnvelope As MsoEnvelope  (read-only)`  
  Represents an email header for a document.
- `CustomProperties As CustomProperties  (read-only)`  
  Returns a CustomProperties object representing the identifier information associated with a worksheet.
- `Protection As Protection  (read-only)`  
  Returns a Protection object that represents the protection options of the worksheet.
- `ListObjects As ListObjects  (read-only)`  
  Returns a collection of ListObject objects on the worksheet. Read-only ListObjects collection.
- `EnableFormatConditionsCalculation As Boolean  (read/write)`  
  Returns or sets if conditional formats occur automatically as needed. Read/write Boolean.
- `PrintedCommentPages As Long  (read-only)`  
  Returns the number of comment pages that will be printed for the current worksheet. Read-only.
- `CommentsThreaded As CommentsThreaded  (read-only)`  
  Returns a CommentsThreaded collection that represents all the top-level/root comments (no replies) for the specified worksheet. Includes legacy and modern comments. Read-only.
- `AutoFilter As AutoFilter  (read-only)`  
  Returns an AutoFilter object if filtering is on. Read-only.
- `Sort As Sort  (read-only)`  
  Returns a Sort object. Read-only.
- `NamedSheetViews As NamedSheetViewCollection  (read-only)`

## Methods (31)

- `Activate()`  
  Makes the current sheet the active sheet.
- `Copy([Before As Variant], [After As Variant])`  
  Copies the sheet to another location in the current workbook or a new workbook.
    - `Before As Variant` (optional): The sheet before which the copied sheet will be placed. You cannot specify _Before_ if you specify _After_.
    - `After As Variant` (optional): The sheet after which the copied sheet will be placed. You cannot specify _After_ if you specify _Before_.
- `Delete()`  
  Deletes the object.
- `Move([Before As Variant], [After As Variant])`  
  Moves the sheet to another location in the workbook.
    - `Before As Variant` (optional): The sheet before which the moved sheet will be placed. You cannot specify _Before_ if you specify _After_.
    - `After As Variant` (optional): The sheet after which the moved sheet will be placed. You cannot specify _After_ if you specify _Before_.
- `PrintPreview([EnableChanges As Variant])`  
  Shows a preview of the object as it would look when printed.
    - `EnableChanges As Variant` (optional): Passes a Boolean value to specify if the user can change the margins and other page setup options available in print preview.
- `Select([Replace As Variant])`  
  Selects the object.
    - `Replace As Variant` (optional): (used only with sheets). True to replace the current selection with the specified object. False to extend the current selection to include any previously selected objects and the specified object.
- `Unprotect([Password As Variant])`  
  Removes protection from a sheet or workbook. This method has no effect if the sheet or workbook isn't protected.
    - `Password As Variant` (optional): A string that denotes the case-sensitive password to use to unprotect the sheet or workbook. If the sheet or workbook isn't protected with a password, this argument is ignored. If you omit this argument for a sheet that's protected with a password, you'll be prompted for the password. If you omit this argument for a workbook that's protected with a password, the method fails.
- `SetBackgroundPicture(Filename As String)`  
  Sets the background graphic for a worksheet.
    - `Filename As String` (required): The name of the graphic file.
- `Calculate()`  
  Calculates all open workbooks, a specific worksheet in a workbook, or a specified range of cells on a worksheet, as shown in the following table.
- `ChartObjects([Index As Variant]) As Object`  
  Returns an object that represents either a single embedded chart (a ChartObject object) or a collection of all the embedded charts (a ChartObjects object) on the sheet.
    - `Index As Variant` (optional): The name or number of the chart. This argument can be an array to specify more than one chart.
- `CheckSpelling([CustomDictionary As Variant], [IgnoreUppercase As Variant], [AlwaysSuggest As Variant], [SpellLang As Variant])`  
  Checks the spelling of an object.
    - `CustomDictionary As Variant` (optional): A string that indicates the file name of the custom dictionary to be examined if the word isn't found in the main dictionary. If this argument is omitted, the currently specified dictionary is used.
    - `IgnoreUppercase As Variant` (optional): True to have Microsoft Excel ignore words that are all uppercase. False to have Excel check words that are all uppercase. If this argument is omitted, the current setting will be used.
    - `AlwaysSuggest As Variant` (optional): True to have Excel display a list of suggested alternate spellings when an incorrect spelling is found. False to have Excel wait for you to input the correct spelling. If this argument is omitted, the current setting will be used.
    - `SpellLang As Variant` (optional): The language of the dictionary being used. Can be one of the MsoLanguageID values.
- `ClearArrows()`  
  Clears the tracer arrows from the worksheet. Tracer arrows are added by using the auditing feature.
- `Evaluate(Name As Variant) As Variant`  
  Converts a Microsoft Excel name to an object or a value.
    - `Name As Variant` (required): The name of the object, using the naming convention of Excel.
- `_Evaluate(Name As Variant) As Variant`
- `ResetAllPageBreaks()`  
  Resets all page breaks on the specified worksheet.
- `OLEObjects([Index As Variant]) As Object`  
  Returns an object that represents either a single OLE object (an OLEObject) or a collection of all OLE objects (an OLEObjects collection) on the chart or sheet. Read-only.
    - `Index As Variant` (optional): The name or number of the OLE object.
- `Paste([Destination As Variant], [Link As Variant])`  
  Pastes the contents of the Clipboard onto the sheet.
    - `Destination As Variant` (optional): A Range object that specifies where the Clipboard contents should be pasted. If this argument is omitted, the current selection is used. This argument can be specified only if the contents of the Clipboard can be pasted into a range. If this argument is specified, the _Link_ argument cannot be used.
    - `Link As Variant` (optional): True to establish a link to the source of the pasted data. If this argument is specified, the _Destination_ argument cannot be used. The default value is False.
- `PivotTables([Index As Variant]) As Object`  
  Returns an object that represents either a single PivotTable report (a PivotTable object) or a collection of all the PivotTable reports (a PivotTables object) on a worksheet. Read-only.
    - `Index As Variant` (optional): The name or number of the report.
- `PivotTableWizard([SourceType As Variant], [SourceData As Variant], [TableDestination As Variant], [TableName As Variant], [RowGrand As Variant], [ColumnGrand As Variant], [SaveData As Variant], [HasAutoFormat As Variant], [AutoPage As Variant], [Reserved As Variant], [BackgroundQuery As Variant], [OptimizeCache As Variant], [PageFieldOrder As Variant], [PageFieldWrapCount As Variant], [ReadData As Variant], [Connection As Variant]) As PivotTable`  
  Creates a new PivotTable report. This method doesn't display the PivotTable Wizard. This method isn't available for OLE DB data sources. Use the Add method to add a PivotTable cache, and then create a PivotTable report based on the cache.
    - `SourceType As Variant` (optional): An XlPivotTableSourceType value that represents the source of the report data. If you specify this argument, you must also specify _SourceData_. If _SourceType_ and _SourceData_ are omitted, Microsoft Excel assumes that the source type is xlDatabase, and the source data comes from the named range Database. If this named range doesn't exist, Excel uses the current region if the current selection is in a range of more than 10 cells that contain data. If this isn't true, this method will fail.
    - `SourceData As Variant` (optional): The data for the new report. Can be a Range object, an array of ranges, or a text constant that represents the name of another report. For an external database, _SourceData_ is an array of strings containing the SQL query string, where each element is up to 255 characters in length. You should use the _Connection_ argument to specify the ODBC connection string. For compatibility with earlier versions of Excel, _SourceData_ can be a two-element array. The first element is the connection string specifying the ODBC source for the data. The second element is the SQL query string used to get the data. If you specify _SourceData_, you must also specify _SourceType_. If the active cell is inside the _SourceData_ range, you must specify _TableDestination_ as well.
    - `TableDestination As Variant` (optional): A Range object specifying where the report should be placed on the worksheet. If this argument is omitted, the report is placed at the active cell.
    - `TableName As Variant` (optional): A string that specifies the name of the new report.
    - `RowGrand As Variant` (optional): True to show grand totals for rows in the report.
    - `ColumnGrand As Variant` (optional): True to show grand totals for columns in the report.
    - `SaveData As Variant` (optional): True to save data with the report. False to save only the report definition.
    - `HasAutoFormat As Variant` (optional): True to have Excel automatically format the report when it's refreshed or when fields are moved.
    - `AutoPage As Variant` (optional): Valid only if _SourceType_ is xlConsolidation. True to have Excel create a page field for the consolidation. If _AutoPage_ is False, you must create the page field or fields.
    - `Reserved As Variant` (optional): Not used by Excel.
    - `BackgroundQuery As Variant` (optional): True to have Excel perform queries for the report asynchronously (in the background). The default value is False.
    - `OptimizeCache As Variant` (optional): True to optimize the PivotTable cache when it's constructed. The default value is False.
    - `PageFieldOrder As Variant` (optional): The order in which page fields are added to the PivotTable report's layout. Can be one of the following XlOrder constants: xlDownThenOver (default) or xlOverThenDown.
    - `PageFieldWrapCount As Variant` (optional): The number of page fields in each column or row in the PivotTable report. The default value is 0 (zero).
    - `ReadData As Variant` (optional): True to create a PivotTable cache that contains all records from the external database; this cache can be very large. If _ReadData_ is False, you can set some of the fields as server-based page fields before the data is actually read.
    - `Connection As Variant` (optional): A string that contains ODBC settings that allow Excel to connect to an ODBC data source. The connection string has the form ODBC;<connection string>. This argument overrides any previous setting for the PivotCache object's Connection property.
- `Scenarios([Index As Variant]) As Object`  
  Returns an object that represents either a single scenario (a Scenario object) or a collection of scenarios (a Scenarios object) on the worksheet.
    - `Index As Variant` (optional): The name or number of the scenario. Use an array to specify more than one scenario.
- `ShowAllData()`  
  Makes all rows of the currently filtered list visible. If AutoFilter is in use, this method changes the arrows to "All."
- `ShowDataForm()`  
  Displays the data form associated with the worksheet.
- `ClearCircles()`  
  Clears circles from invalid entries on the worksheet.
- `CircleInvalid()`  
  Circles invalid entries on the worksheet.
- `PasteSpecial([Format As Variant], [Link As Variant], [DisplayAsIcon As Variant], [IconFileName As Variant], [IconIndex As Variant], [IconLabel As Variant], [NoHTMLFormatting As Variant])`  
  Pastes the contents of the Clipboard onto the sheet, using a specified format. Use this method to paste data from other applications or to paste data in a specific format.
    - `Format As Variant` (optional): A string that specifies the Clipboard format of the data.
    - `Link As Variant` (optional): True to establish a link to the source of the pasted data. If the source data isn't suitable for linking or the source application doesn't support linking, this parameter is ignored. The default value is False.
    - `DisplayAsIcon As Variant` (optional): True to display the pasted data as an icon. The default value is False.
    - `IconFileName As Variant` (optional): The name of the file that contains the icon to use if _DisplayAsIcon_ is True.
    - `IconIndex As Variant` (optional): The index number of the icon within the icon file.
    - `IconLabel As Variant` (optional): The text label of the icon.
    - `NoHTMLFormatting As Variant` (optional): True to remove all formatting, hyperlinks, and images from HTML. False to paste HTML as is. The default value is False.
- `Protect([Password As Variant], [DrawingObjects As Variant], [Contents As Variant], [Scenarios As Variant], [UserInterfaceOnly As Variant], [AllowFormattingCells As Variant], [AllowFormattingColumns As Variant], [AllowFormattingRows As Variant], [AllowInsertingColumns As Variant], [AllowInsertingRows As Variant], [AllowInsertingHyperlinks As Variant], [AllowDeletingColumns As Variant], [AllowDeletingRows As Variant], [AllowSorting As Variant], [AllowFiltering As Variant], [AllowUsingPivotTables As Variant])`  
  Protects a worksheet so that it cannot be modified.
    - `Password As Variant` (optional): A string that specifies a case-sensitive password for the worksheet or workbook. If this argument is omitted, you can unprotect the worksheet or workbook without using a password. Otherwise, you must specify the password to unprotect the worksheet or workbook. If you forget the password, you cannot unprotect the worksheet or workbook. Use strong passwords that combine uppercase and lowercase letters, numbers, and symbols. Weak passwords don't mix these elements. Strong password: Y6dh!et5. Weak password: House27. Passwords should be 8 or more characters in length. A pass phrase that uses 14 or more characters is better. It's critical that you remember your password. If you forget your password, Microsoft cannot retrieve it. Store the passwords that you write down in a secure place away from the information that they help protect.
    - `DrawingObjects As Variant` (optional): True to protect shapes. The default value is True.
    - `Contents As Variant` (optional): True to protect contents. For a chart, this protects the entire chart. For a worksheet, this protects the locked cells. The default value is True.
    - `Scenarios As Variant` (optional): True to protect scenarios. This argument is valid only for worksheets. The default value is True.
    - `UserInterfaceOnly As Variant` (optional): True to protect the user interface, but not macros. If this argument is omitted, protection applies both to macros and to the user interface.
    - `AllowFormattingCells As Variant` (optional): True allows the user to format any cell on a protected worksheet. The default value is False.
    - `AllowFormattingColumns As Variant` (optional): True allows the user to format any column on a protected worksheet. The default value is False.
    - `AllowFormattingRows As Variant` (optional): True allows the user to format any row on a protected worksheet. The default value is False.
    - `AllowInsertingColumns As Variant` (optional): True allows the user to insert columns on the protected worksheet. The default value is False.
    - `AllowInsertingRows As Variant` (optional): True allows the user to insert rows on the protected worksheet. The default value is False.
    - `AllowInsertingHyperlinks As Variant` (optional): True allows the user to insert hyperlinks on the protected worksheet. The default value is False.
    - `AllowDeletingColumns As Variant` (optional): True allows the user to delete columns on the protected worksheet, where every cell in the column to be deleted is unlocked. The default value is False.
    - `AllowDeletingRows As Variant` (optional): True allows the user to delete rows on the protected worksheet, where every cell in the row to be deleted is unlocked. The default value is False.
    - `AllowSorting As Variant` (optional): True allows the user to sort on the protected worksheet. Every cell in the sort range must be unlocked or unprotected. The default value is False.
    - `AllowFiltering As Variant` (optional): True allows the user to set filters on the protected worksheet. Users can change filter criteria but can not enable or disable an auto filter. Users can set filters on an existing auto filter. The default value is False.
    - `AllowUsingPivotTables As Variant` (optional): True allows the user to use PivotTable reports on the protected worksheet. The default value is False.
- `XmlDataQuery(XPath As String, [SelectionNamespaces As Variant], [Map As Variant]) As Range`  
  Returns a Range object that represents the cells mapped to a particular XPath. Returns Nothing if the specified XPath has not been mapped to the worksheet, or if the mapped range is empty.
    - `XPath As String` (required): The XPath to query for.
    - `SelectionNamespaces As Variant` (optional): A space-delimited String that contains the namespaces referenced in the XPath parameter. A run-time error is generated if one of the specified namespaces cannot be resolved.
    - `Map As Variant` (optional): Specify an XmlMap if you want to query for the XPath within a specific map.
- `XmlMapQuery(XPath As String, [SelectionNamespaces As Variant], [Map As Variant]) As Range`  
  Returns a Range object that represents the cells mapped to a particular XPath. Returns Nothing if the specified XPath has not been mapped to the worksheet.
    - `XPath As String` (required): The XPath to query for.
    - `SelectionNamespaces As Variant` (optional): A space-delimited String that contains the namespaces referenced in the XPath parameter. A run-time error is generated if one of the specified namespaces cannot be resolved.
    - `Map As Variant` (optional): Specify an XML map if you want to query for the XPath within a specific map.
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
- `ExportAsFixedFormat(Type As XlFixedFormatType, [Filename As Variant], [Quality As Variant], [IncludeDocProperties As Variant], [IgnorePrintAreas As Variant], [From As Variant], [To As Variant], [OpenAfterPublish As Variant], [FixedFormatExtClassPtr As Variant], [WorkIdentity As Variant])`  
  Exports to a file of the specified format.
    - `Type As XlFixedFormatType` (required): The type of file format to export to.
    - `Filename As Variant` (optional): The file name of the file to be saved. You can include a full path, or Excel saves the file in the current folder.
    - `Quality As Variant` (optional): Optional XlFixedFormatQuality. Specifies the quality of the published file.
    - `IncludeDocProperties As Variant` (optional): True to include the document properties; otherwise, False.
    - `IgnorePrintAreas As Variant` (optional): True to ignore any print areas set when publishing; otherwise, False.
    - `From As Variant` (optional): The number of the page at which to start publishing. If this argument is omitted, publishing starts at the beginning.
    - `To As Variant` (optional): The number of the last page to publish. If this argument is omitted, publishing ends with the last page.
    - `OpenAfterPublish As Variant` (optional): True to display the file in the viewer after it is published; otherwise, False.
    - `FixedFormatExtClassPtr As Variant` (optional): Pointer to the FixedFormatExt class.
- `SaveAs(Filename As String, [FileFormat As Variant], [Password As Variant], [WriteResPassword As Variant], [ReadOnlyRecommended As Variant], [CreateBackup As Variant], [AddToMru As Variant], [TextCodepage As Variant], [TextVisualLayout As Variant], [Local As Variant])`  
  Saves changes to the chart or worksheet in a different file.
    - `Filename As String` (required): Variant. A string that indicates the name of the file to be saved. You can include a full path; if you don't, Microsoft Excel saves the file in the current folder.
    - `FileFormat As Variant` (optional): The file format to use when you save the file. For a list of valid choices, see the XlFileFormat enumeration. For an existing file, the default format is the last file format specified; for a new file, the default is the format of the version of Excel being used.
    - `Password As Variant` (optional): A case-sensitive string (no more than 15 characters) that indicates the protection password to be given to the file.
    - `WriteResPassword As Variant` (optional): A string that indicates the write-reservation password for this file. If a file is saved with the password and the password isn't supplied when the file is opened, the file is opened as read-only.
    - `ReadOnlyRecommended As Variant` (optional): True to display a message when the file is opened, recommending that the file be opened as read-only.
    - `CreateBackup As Variant` (optional): True to create a backup file.
    - `AddToMru As Variant` (optional): True to add this workbook to the list of recently used files. The default value is False.
    - `TextCodepage As Variant` (optional): Not used in U.S. English Microsoft Excel.
    - `TextVisualLayout As Variant` (optional): Not used in U.S. English Microsoft Excel.
    - `Local As Variant` (optional): True saves files against the language of Excel (including control panel settings). False (default) saves files against the language of Visual Basic for Applications (VBA) (which is typically US English unless the VBA project where Workbooks.Open is run from is an old internationalized XL5/95 VBA project).

## Events (17)

- `SelectionChange(Target As Range)`  
  Occurs when the selection changes on a worksheet.
    - `Target As Range` (required): The new selected range.
- `BeforeDoubleClick(Target As Range, Cancel As Boolean)`  
  Occurs when a worksheet is double-clicked, before the default double-click action.
    - `Target As Range` (required): The cell nearest to the mouse pointer when the double-click occurs.
    - `Cancel As Boolean` (required): False when the event occurs. If the event procedure sets this argument to True, the default double-click action isn't performed when the procedure is finished.
- `BeforeRightClick(Target As Range, Cancel As Boolean)`  
  Occurs when a worksheet is right-clicked, before the default right-click action.
    - `Target As Range` (required): The cell nearest to the mouse pointer when the right-click occurs.
    - `Cancel As Boolean` (required): False when the event occurs. If the event procedure sets this argument to True, the default right-click action doesn't occur when the procedure is finished.
- `Activate()`  
  Occurs when a workbook, worksheet, chart sheet, or embedded chart is activated.
- `Deactivate()`  
  Occurs when the chart, worksheet, or workbook is deactivated.
- `Calculate()`  
  Occurs after the worksheet is recalculated for the Worksheet object.
- `Change(Target As Range)`  
  Occurs when cells on the worksheet are changed by the user or by an external link.
    - `Target As Range` (required): The changed range. Can be more than one cell.
- `FollowHyperlink(Target As Hyperlink)`  
  Occurs when you choose any hyperlink on a worksheet. For application- and workbook-level events, see the Application.SheetFollowHyperlink event and Workbook.SheetFollowHyperlink event.
    - `Target As Hyperlink` (required): A Hyperlink object that represents the destination of the hyperlink.
- `PivotTableUpdate(Target As PivotTable)`  
  Occurs after a PivotTable report is updated on a worksheet.
    - `Target As PivotTable` (required): The selected PivotTable report.
- `PivotTableAfterValueChange(TargetPivotTable As PivotTable, TargetRange As Range)`  
  Occurs after a cell or range of cells inside a PivotTable are edited or recalculated (for cells that contain formulas).
    - `TargetPivotTable As PivotTable` (required): The PivotTable that contains the edited or recalculated cells.
    - `TargetRange As Range` (required): The range that contains all the edited or recalculated cells.
- `PivotTableBeforeAllocateChanges(TargetPivotTable As PivotTable, ValueChangeStart As Long, ValueChangeEnd As Long, Cancel As Boolean)`  
  Occurs before changes are applied to a PivotTable.
    - `TargetPivotTable As PivotTable` (required): The PivotTable that contains the changes to apply.
    - `ValueChangeStart As Long` (required): The index to the first change in the associated PivotTableChangeList collection. The index is specified by the Order property of the ValueChange object in the PivotTableChangeList collection.
    - `ValueChangeEnd As Long` (required): The index to the last change in the associated PivotTableChangeList collection. The index is specified by the Order property of the ValueChange object in the PivotTableChangeList collection.
    - `Cancel As Boolean` (required): False when the event occurs. If the event procedure sets this argument to True, the changes are not applied to the PivotTable and all edits are lost.
- `PivotTableBeforeCommitChanges(TargetPivotTable As PivotTable, ValueChangeStart As Long, ValueChangeEnd As Long, Cancel As Boolean)`  
  Occurs before changes are committed against the OLAP data source for a PivotTable.
    - `TargetPivotTable As PivotTable` (required): The PivotTable that contains the changes to commit.
    - `ValueChangeStart As Long` (required): The index to the first change in the associated PivotTableChangeList object. The index is specified by the Order property of the ValueChange object in the PivotTableChangeList collection.
    - `ValueChangeEnd As Long` (required): The index to the last change in the associated PivotTableChangeList object. The index is specified by the Order property of the ValueChange object in the PivotTableChangeList collection.
    - `Cancel As Boolean` (required): False when the event occurs. If the event procedure sets this argument to True, the changes are not committed against the OLAP data source of the PivotTable.
- `PivotTableBeforeDiscardChanges(TargetPivotTable As PivotTable, ValueChangeStart As Long, ValueChangeEnd As Long)`  
  Occurs before changes to a PivotTable are discarded.
    - `TargetPivotTable As PivotTable` (required): The PivotTable that contains the changes to discard.
    - `ValueChangeStart As Long` (required): The index to the first change in the associated PivotTableChangeList object. The index is specified by the Order property of the ValueChange object in the PivotTableChangeList collection.
    - `ValueChangeEnd As Long` (required): The index to the last change in the associated PivotTableChangeList object. The index is specified by the Order property of the ValueChange object in the PivotTableChangeList collection.
- `PivotTableChangeSync(Target As PivotTable)`  
  Occurs after changes to a PivotTable.
    - `Target As PivotTable` (required): The PivotTable that was changed.
- `LensGalleryRenderComplete()`  
  Occurs when a callout gallery's icons (dynamic and static) have completed rendering.
- `TableUpdate(Target As TableObject)`  
  Occurs after a query table connected to the data model is updated on a worksheet.
    - `Target As TableObject` (required): The query table with data from the data model.
- `BeforeDelete()`  
  Occurs before the worksheet is deleted.
