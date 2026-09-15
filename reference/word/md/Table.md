# Table

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020951-0000-0000-C000-000000000046}  

Represents a single table. The Table object is a member of the Tables collection. The Tables collection includes all the tables in the specified selection, range, or document.

**Remarks:** Use Tables (Index), where Index is the index number, to return a single Table object. The index number represents the position of the table in the selection, range, or document. The following example converts the first table in the active document to text. Use the Add method to add a table at the specified range. The following example adds a 3x4 table at the beginning of the active document.

## Properties (31)

- `Range As Range  (read-only)`  
  Returns a Range object that represents the portion of a document that is contained within the specified table.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Table object.
- `Columns As Columns  (read-only)`  
  Returns a Columns collection that represents all the table columns in the table. Read-only.
- `Rows As Rows  (read-only)`  
  Returns a Rows collection that represents all the table rows within a table. Read-only.
- `Borders As Borders  (read/write)`  
  Returns a Borders collection that represents all the borders for the specified object.
- `Shading As Shading  (read-only)`  
  Returns a Shading object that refers to the shading formatting for the specified object.
- `Uniform As Boolean  (read-only)`  
  True if all the rows in a table have the same number of columns. Read-only Boolean.
- `AutoFormatType As Long  (read-only)`  
  Returns the type of automatic formatting that's been applied to the specified table. Read-only Long.
- `Tables As Tables  (read-only)`  
  Returns a Tables collection that represents all the tables nested within the specified table. Read-only.
- `NestingLevel As Long  (read-only)`  
  Returns the nesting level of the specified table. Read-only Long.
- `AllowAutoFit As Boolean  (read/write)`  
  Allows Microsoft Word to automatically resize cells in a table to fit their contents. Read/write Boolean.
- `PreferredWidth As Single  (read/write)`  
  Returns or sets the preferred width (in points or as a percentage of the window width) for the specified table. Read/write Single.
- `PreferredWidthType As WdPreferredWidthType  (read/write)`  
  Returns or sets the preferred unit of measurement to use for the width of the specified table. Read/write WdPreferredWidthType.
- `TopPadding As Single  (read/write)`  
  Returns or sets the amount of space (in points) to add above the contents of all the cells in a table. Read/write Single.
- `BottomPadding As Single  (read/write)`  
  Returns or sets the amount of space (in points) to add below the contents of a single cell or all the cells in a table. Read/write Single.
- `LeftPadding As Single  (read/write)`  
  Returns or sets the amount of space (in points) to add to the left of the contents of all the cells in a table. Read/write Single.
- `RightPadding As Single  (read/write)`  
  Returns or sets the amount of space (in points) to add to the right of the contents of all the cells in a table. Read/write Single.
- `Spacing As Single  (read/write)`  
  Returns or sets the spacing (in points) between the cells in a table. Read/write Single.
- `TableDirection As WdTableDirection  (read/write)`  
  Returns or sets the direction in which Microsoft Word orders cells in the specified table. Read/write WdTableDirection.
- `ID As String  (read/write)`  
  Returns or sets the identifying label for the specified table when the document is saved as a webpage. Read/write String.
- `Style As Variant  (read/write)`  
  Returns or sets the style for the specified table. Read/write Variant.
- `ApplyStyleHeadingRows As Boolean  (read/write)`  
  True for Microsoft Word to apply heading-row formatting to the first row of the selected table. Read/write Boolean.
- `ApplyStyleLastRow As Boolean  (read/write)`  
  True for Microsoft Word to apply last-row formatting to the last row of the specified table. Read/write Boolean.
- `ApplyStyleFirstColumn As Boolean  (read/write)`  
  True for Microsoft Word to apply first-column formatting to the first column of the specified table. Read/write Boolean.
- `ApplyStyleLastColumn As Boolean  (read/write)`  
  True for Microsoft Word to apply last-column formatting to the last column of the specified table. Read/write Boolean.
- `ApplyStyleRowBands As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether to apply style bands to the rows in a table if an applied preset table style provides style banding for rows. Read/write.
- `ApplyStyleColumnBands As Boolean  (read/write)`  
  Returns or sets a Boolean that represents whether to apply style bands to the columns in a table if an applied preset table style provides style banding for columns. Read/write.
- `Title As String  (read/write)`  
  Returns or sets a String that contains a title for the specified table. Read/write.
- `Descr As String  (read/write)`  
  Returns or sets a String that contains a description for the specified table. Read/write.

## Methods (12)

- `Select()`  
  Selects the specified table.
- `Delete()`  
  Deletes the specified table.
- `SortAscending()`  
  Sorts paragraphs or table rows in ascending alphanumeric order.
- `SortDescending()`  
  Sorts table rows in descending alphanumeric order.
- `AutoFormat([Format As Variant], [ApplyBorders As Variant], [ApplyShading As Variant], [ApplyFont As Variant], [ApplyColor As Variant], [ApplyHeadingRows As Variant], [ApplyLastRow As Variant], [ApplyFirstColumn As Variant], [ApplyLastColumn As Variant], [AutoFit As Variant])`  
  Applies a predefined look to a table.
    - `Format As Variant` (optional): The format to apply. This parameter can be a WdTableFormat constant, a WdTableFormatApply constant, or a TableStyle object.
    - `ApplyBorders As Variant` (optional): True to apply the border properties of the specified format. The default value is True.
    - `ApplyShading As Variant` (optional): True to apply the shading properties of the specified format. The default value is True.
    - `ApplyFont As Variant` (optional): True to apply the font properties of the specified format. The default value is True.
    - `ApplyColor As Variant` (optional): True to apply the color properties of the specified format. The default value is True.
    - `ApplyHeadingRows As Variant` (optional): True to apply the heading-row properties of the specified format. The default value is True.
    - `ApplyLastRow As Variant` (optional): True to apply the last-row properties of the specified format. The default value is False.
    - `ApplyFirstColumn As Variant` (optional): True to apply the first-column properties of the specified format. The default value is True.
    - `ApplyLastColumn As Variant` (optional): True to apply the last-column properties of the specified format. The default value is False.
    - `AutoFit As Variant` (optional): True to decrease the width of the table columns as much as possible without changing the way text wraps in the cells. The default value is True.
- `UpdateAutoFormat()`  
  Updates the table with the characteristics of a predefined table format.
- `Cell(Row As Long, Column As Long) As Cell`  
  Returns a Cell object that represents a cell in a table.
    - `Row As Long` (required): The number of the row in the table to return. Can be an integer between 1 and the number of rows in the table.
    - `Column As Long` (required): The number of the cell in the table to return. Can be an integer between 1 and the number of columns in the table.
- `Split(BeforeRow As Variant) As Table`  
  Inserts an empty paragraph immediately above the specified row in the table, and returns a Table object that contains both the specified row and the rows that follow it.
    - `BeforeRow As Variant` (required): The row that the table is to be split before. Can be a row number or a Row object.
- `ConvertToText([Separator As Variant], [NestedTables As Variant]) As Range`  
  Converts a table to text and returns a Range object that represents the delimited text.
    - `Separator As Variant` (optional): The character that delimits the converted columns (paragraph marks delimit the converted rows). Can be any WdTableFieldSeparator constants.
    - `NestedTables As Variant` (optional): True if nested tables are converted to text. This argument is ignored if Separator is not wdSeparateByParagraphs. The default value is True.
- `AutoFitBehavior(Behavior As WdAutoFitBehavior)`  
  Determines how Microsoft Word resizes a table when the AutoFit feature is used.
    - `Behavior As WdAutoFitBehavior` (required): How Word resizes the specified table with the AutoFit feature is used.
- `Sort([ExcludeHeader As Variant], [FieldNumber As Variant], [SortFieldType As Variant], [SortOrder As Variant], [FieldNumber2 As Variant], [SortFieldType2 As Variant], [SortOrder2 As Variant], [FieldNumber3 As Variant], [SortFieldType3 As Variant], [SortOrder3 As Variant], [CaseSensitive As Variant], [BidiSort As Variant], [IgnoreThe As Variant], [IgnoreKashida As Variant], [IgnoreDiacritics As Variant], [IgnoreHe As Variant], [LanguageID As Variant])`  
  Sorts the specified table.
    - `ExcludeHeader As Variant` (optional): True to exclude the first row from the sort operation. The default value is False.
    - `FieldNumber As Variant` (optional): The first field by which to sort. Microsoft Word sorts by FieldNumber, then by FieldNumber2, and then by FieldNumber3.
    - `SortFieldType As Variant` (optional): The sort type for FieldNumber. Can be one of the WdSortFieldType constants. Some of these constants may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed. The default value is wdSortFieldAlphanumeric.
    - `SortOrder As Variant` (optional): The sorting order to use when sorting FieldNumber. Can be a WdSortOrder constant.
    - `FieldNumber2 As Variant` (optional): The second field by which to sort.
    - `SortFieldType2 As Variant` (optional): The sort type for FieldNumber2. Can be one of the WdSortFieldType constants. Some of these constants may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed. The default value is wdSortFieldAlphanumeric.
    - `SortOrder2 As Variant` (optional): The sorting order to use when sorting FieldNumber2. Can be one WdSortOrder constant.
    - `FieldNumber3 As Variant` (optional): The third field by which to sort.
    - `SortFieldType3 As Variant` (optional): The sort type for FieldNumber3. Can be one of the WdSortFieldType constants. Some of these constants may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed. The default value is wdSortFieldAlphanumeric.
    - `SortOrder3 As Variant` (optional): The sorting order to use when sorting FieldNumber3. Can be one WdSortOrder constant.
    - `CaseSensitive As Variant` (optional): True to sort with case sensitivity. The default value is False.
    - `BidiSort As Variant` (optional): True to sort based on right-to-left language rules. This argument may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.
    - `IgnoreThe As Variant` (optional): True to ignore the Arabic character alef lam when sorting right-to-left language text. This argument may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.
    - `IgnoreKashida As Variant` (optional): True to ignore kashidas when sorting right-to-left language text. This argument may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.
    - `IgnoreDiacritics As Variant` (optional): True to ignore bidirectional control characters when sorting right-to-left language text. This argument may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.
    - `IgnoreHe As Variant` (optional): True to ignore the Hebrew character he when sorting right-to-left language text. This argument may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.
    - `LanguageID As Variant` (optional): Specifies the sorting language. Can be one of the WdLanguageID constants. Refer to the Object Browser for a list of the WdLanguageID constants.
- `ApplyStyleDirectFormatting(StyleName As String)`  
  Applies the specified style but maintains any formatting that a user directly applies.
    - `StyleName As String` (required): The name of the style to apply.
