# ListObjects

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024470-0000-0000-C000-000000000046}  

A collection of all the ListObject objects on a worksheet. Each ListObject object represents a table on the worksheet.

**Remarks:** Use the ListObjects property of the Worksheet object to return the ListObjects collection.

**Example:**

```vba
Set myWorksheetLists = Worksheets(1).ListObjects
```

## Properties (7)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `_Default As ListObject  (read-only)`
- `_NewEnum As IUnknown  (read-only)`
- `Item As ListObject  (read-only)`  
  Returns a single object from a collection.
- `Count As Long  (read-only)`  
  Returns an Integer value that represents the number of objects in the collection.

## Methods (1)

- `Add([SourceType As XlListObjectSourceType], [Source As Variant], [LinkSource As Variant], [XlListObjectHasHeaders As XlYesNoGuess], [Destination As Variant], [TableStyleName As Variant]) As ListObject`  
  Creates a new list object.
    - `SourceType As XlListObjectSourceType` (optional): Indicates the kind of source for the query.
    - `Source As Variant` (optional): When _SourceType_ = xlSrcRange: A Range object representing the data source. If omitted, the _Source_ will default to the range returned by list range detection code. When _SourceType_ = xlSrcExternal: An array of String values specifying a connection to the source, containing the following elements:<ul><li>0 - URL to SharePoint site</li><li>1 - ListName</li><li>2 - ViewGUID</li></ul> When _SourceType_ = xlSrcQuery: Accepts ODBC or OLEDB connection string (this also works with xlSrcExternal). When _SourceType_ = xlSrcModel: Accepts a WorkbookConnection object (see TableObject for example).
    - `LinkSource As Variant` (optional): Indicates whether an external data source is to be linked to the ListObject object. If _SourceType_ is xlSrcExternal, the default is True. Invalid if _SourceType_ is xlSrcRange, and will return an error if not omitted.
    - `XlListObjectHasHeaders As XlYesNoGuess` (optional): An XlYesNoGuess constant that indicates whether the data being imported has column labels. If the _Source_ does not contain headers, Excel will automatically generate headers. Default value: xlGuess.
    - `Destination As Variant` (optional): A Range object specifying a single-cell reference as the destination for the top-left corner of the new list object. If the Range object refers to more than one cell, an error is generated. The _Destination_ argument must be specified when _SourceType_ is set to xlSrcExternal. The _Destination_ argument is ignored if _SourceType_ is set to xlSrcRange. The destination range must be on the worksheet that contains the ListObjects collection specified by _expression_. New columns will be inserted at the _Destination_ to fit the new list. Therefore, existing data will not be overwritten.
    - `TableStyleName As Variant` (optional): The name of a TableStyle; for example "TableStyleLight1".
