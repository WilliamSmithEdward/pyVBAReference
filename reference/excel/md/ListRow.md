# ListRow

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024475-0000-0000-C000-000000000046}  

Represents a row in a table. The ListRow object is a member of the ListRows collection.

**Remarks:** The ListRows collection contains all the rows in a list object. Use the ListRows property of the ListObject object to return a ListRows collection.

**Example:**

```vba
Dim wrksht As Worksheet
Dim oListRow As ListRow

Set wrksht = ActiveWorkbook.Worksheets("Sheet1")
Set oListRow = wrksht.ListObjects(1).ListRows.Add
```

## Properties (5)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Index As Long  (read-only)`  
  Returns a Long value that represents the index number of the ListRow object within the ListRows collection.
- `Range As Range  (read-only)`  
  Returns a Range object that represents the range to which the specified list object in the list applies.

## Methods (1)

- `Delete()`  
  Deletes the cells of the list row and shifts upward any remaining cells below the deleted row. You can delete rows in the list even when the list is linked to a SharePoint site. The list on the SharePoint site will not be updated, however, until you synchronize your changes.
