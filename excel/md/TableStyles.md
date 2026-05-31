# TableStyles

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244A8-0000-0000-C000-000000000046}  

Represents styles that can be applied to a table.

**Remarks:** Table styles provide a way to format an entire table or PivotTable. Table styles replace the existing auto format feature for formatting an entire table. Table styles differ from auto format in the following ways: - You can create and reuse custom table styles. - Table styles work with themes. - Changing the document theme color scheme and/or font scheme will change the look of the built-in table styles. - Table styles can reapply styles to objects such as PivotTables and tables as the object changes. The tables will remember the style applied to them and will re-display appropriately when cells are added, removed, hidden, and shown. - Table styles have a visual user interface in the ribbon.

## Properties (6)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Count As Long  (read-only)`  
  Returns the number of objects in the collection. Read-only Long.
- `_Default As TableStyle  (read-only)`
- `_NewEnum As IUnknown  (read-only)`

## Methods (2)

- `Add(TableStyleName As String) As TableStyle`  
  Creates a new TableStyle object and adds it to the collection.
    - `TableStyleName As String` (required): The name of the table style.
- `Item(Index As Variant) As TableStyle`  
  Returns a single object from a collection.
    - `Index As Variant` (required): The name or index number for the object.
