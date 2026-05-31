# ListObject

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024471-0000-0000-C000-000000000046}  

Represents a list object in the ListObjects collection.

**Remarks:** The ListObject object is a member of the ListObjects collection. The ListObjects collection contains all the list objects on a worksheet.

**Example:**

```vba
Dim wrksht As Worksheet
Dim oListCol As ListRow

Set wrksht = ActiveWorkbook.Worksheets("Sheet1")
Set oListCol = wrksht.ListObjects(1).ListRows.Add
```

## Properties (35)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `_Default As String  (read-only)`
- `Active As Boolean  (read-only)`  
  Returns a Boolean value indicating whether a ListObject object on a worksheet is active-that is, whether the active cell is inside the range of the ListObject object. Read-only Boolean.
- `DataBodyRange As Range  (read-only)`  
  Returns a Range object that represents the range of values, excluding the header row, in a table. Read-only.
- `DisplayRightToLeft As Boolean  (read-only)`  
  True if the specified ListObject is displayed from right to left instead of from left to right. False if the object is displayed from left to right. Read-only Boolean.
- `HeaderRowRange As Range  (read-only)`  
  Returns a Range object that represents the range of the header row for a list. Read-only Range.
- `InsertRowRange As Range  (read-only)`  
  Returns a Range object representing the Insert row, if any, of a specified ListObject object. Read-only Range.
- `ListColumns As ListColumns  (read-only)`  
  Returns a ListColumns collection that represents all the columns in a ListObject object. Read-only.
- `ListRows As ListRows  (read-only)`  
  Returns a ListRows object that represents all the rows of data in the ListObject object. Read-only.
- `Name As String  (read/write)`  
  Returns or sets a String value that represents the name of the ListObject object.
- `QueryTable As QueryTable  (read-only)`  
  Returns the QueryTable object that provides a link for the ListObject object to the list server. Read-only.
- `Range As Range  (read-only)`  
  Returns a Range object that represents the range to which the specified list object in the list applies.
- `ShowAutoFilter As Boolean  (read/write)`  
  Returns Boolean to indicate whether the AutoFilter will be displayed. Read/write Boolean.
- `ShowTotals As Boolean  (read/write)`  
  Gets or sets a Boolean to indicate whether the Total row is visible. Read/write Boolean.
- `SourceType As XlListObjectSourceType  (read-only)`  
  Returns an XlListObjectSourceType value that represents the current source of the list.
- `TotalsRowRange As Range  (read-only)`  
  Returns a Range object representing the Total row, if any, from a specified ListObject object. Read-only.
- `SharePointURL As String  (read-only)`  
  Returns a String representing the URL of the SharePoint list for a given ListObject object. Read-only String.
- `XmlMap As XmlMap  (read-only)`  
  Returns an XmlMap object that represents the schema map used for the specified table. Read-only.
- `DisplayName As String  (read/write)`  
  Returns or sets the display name for the specified ListObject object. Read/write String.
- `ShowHeaders As Boolean  (read/write)`  
  Returns or sets if the header information should be displayed for the specified ListObject object. Read/write Boolean.
- `TableStyle As Variant  (read/write)`  
  Gets or sets the table style for the specified ListObject object. Read/write Variant.
- `ShowTableStyleFirstColumn As Boolean  (read/write)`  
  Returns or sets if the first column is formatted for the specified ListObject object. Read/write Boolean.
- `ShowTableStyleLastColumn As Boolean  (read/write)`  
  Returns or sets if the last column is displayed for the specified ListObject object. Read/write Boolean.
- `ShowTableStyleRowStripes As Boolean  (read/write)`  
  Returns or sets if the Row Stripes table style is used for the specified ListObject object. Read/write Boolean.
- `ShowTableStyleColumnStripes As Boolean  (read/write)`  
  Returns or sets if the Column Stripes table style is used for the specified ListObject object. Read/write Boolean.
- `Comment As String  (read/write)`  
  Returns or sets the comment associated with the list object. Read/write String.
- `AlternativeText As String  (read/write)`  
  Returns or sets the descriptive (alternative) text string for the specified table. Read/write.
- `Summary As String  (read/write)`  
  Returns or sets the description associated with the alternative text string for the specified table. Read/write.
- `TableObject As TableObject  (read-only)`  
  Returns a TableObject object. Read-only.
- `Slicers As Slicers  (read-only)`  
  Returns a list of the table slicers associated with a ListObject. Read-only.
- `ShowAutoFilterDropDown As Boolean  (read/write)`  
  True when the AutoFilter drop-down for the ListObject object is displayed. Read/write Boolean.
- `AutoFilter As AutoFilter  (read-only)`  
  Filters a table using the AutoFilter feature. Read-only.
- `Sort As Sort  (read-only)`  
  Gets or sets the sort column or columns and sort order for the ListObject collection.

## Methods (7)

- `Delete()`  
  Deletes the ListObject object and clears the cell data from the worksheet.
- `Publish(Target As Variant, LinkSource As Boolean) As String`  
  Publishes the ListObject object to a server that is running Microsoft SharePoint Foundation.
    - `Target As Variant` (required): Contains an array of String values, as described in the Remarks section.
    - `LinkSource As Boolean` (required): See the Remarks section.
- `Refresh()`  
  Retrieves the current data and schema for the list from the server that is running Microsoft SharePoint Foundation. This method can be used only with lists that are linked to a SharePoint site. If the SharePoint site is not available, calling this method returns an error.
- `Unlink()`  
  Removes the link to a Microsoft SharePoint Foundation site from a list. Returns Nothing.
- `Unlist()`  
  Removes the list functionality from a ListObject object. After you use this method, the range of cells that made up the list will be a regular range of data.
- `Resize(Range As Range)`  
  The Resize method allows a ListObject object to be resized over a new range. No cells are inserted or moved.
    - `Range As Range` (required): The new range.
- `ExportToVisio()`  
  Exports a ListObject object to Visio.
