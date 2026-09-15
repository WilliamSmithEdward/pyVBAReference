# Columns

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {0002094B-0000-0000-C000-000000000046}  

A collection of Column objects that represent the columns in a table.

**Remarks:** Use the Columns property of a Range, Selection, or Table object to return a Columns collection. The following example displays the number of Column objects in the Columns collection for the first table in the active document. The following example creates a table with six columns and three rows and then formats each column with a progressively larger (darker) shading percentage. Use the Add method to add a column to a table. The following example adds a column to the first table in the active document, and then it makes the column widths equal. Use Columns (Index), where Index is the index number, to return a single Column object. The index number represents the position of the column in the Columns collection (counting from left to right). The following example selects the first column in the first table.

## Properties (13)

- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of columns in the collection. Read-only.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Columns object.
- `First As Column  (read-only)`  
  Returns a Column object that represents the first item in the Columns collection.
- `Last As Column  (read-only)`  
  Returns a Column object that represents the last column in a table.
- `Width As Single  (read/write)`  
  Returns or sets the width of the specified columns, in points. Read/write Long.
- `Borders As Borders  (read/write)`  
  Returns a Borders collection that represents all the borders for the specified columns.
- `Shading As Shading  (read-only)`  
  Returns a Shading object that refers to the shading formatting for the specified table columns.
- `NestingLevel As Long  (read-only)`  
  Returns the nesting level of the specified columns. Read-only Long.
- `PreferredWidth As Single  (read/write)`  
  Returns or sets the preferred width (in points or as a percentage of the window width) for the specified columns. Read/write Single.
- `PreferredWidthType As WdPreferredWidthType  (read/write)`  
  Returns or sets the preferred unit of measurement to use for the width of the specified cells, columns, or table. Read/write WdPreferredWidthType.

## Methods (7)

- `Item(Index As Long) As Column`  
  Returns an individual Column object in a collection.
    - `Index As Long` (required): The individual object to be returned. Can be a Long indicating the ordinal position of the individual object.
- `Add([BeforeColumn As Variant]) As Column`  
  Returns a Column object that represents a column added to a table.
    - `BeforeColumn As Variant` (optional): A Column object that represents the column that will appear immediately to the right of the new column.
- `Select()`  
  Selects the specified table columns.
- `Delete()`  
  Deletes the specified columns.
- `SetWidth(ColumnWidth As Single, RulerStyle As WdRulerStyle)`  
  Sets the width of columns in a table.
    - `ColumnWidth As Single` (required): The width of the specified column or columns, in points.
    - `RulerStyle As WdRulerStyle` (required): Controls the way Word adjusts cell widths.
- `AutoFit()`  
  Changes the width of a table column to accommodate the width of the text without changing the way text wraps in the cells.
- `DistributeWidth()`  
  Adjusts the width of the specified columns so that they are equal.
