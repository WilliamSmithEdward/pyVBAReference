# TableStyleElements

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244A6-0000-0000-C000-000000000046}  

Represents table style elements.

**Remarks:** Table styles define formatting for one or all of the elements of a table, PivotTable, or slicer. For example, a header row, last column, or total row are elements of a table. A table style can define that the fill color of the header row is blue and that the last column is red. A table style element in a table can have formatting specified in a table style applied to the element. The XlTableStyleElementType enumeration contains the types of table style elements that are available for use.

## Properties (6)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Count As Long  (read-only)`  
  Returns the number of objects in the collection. Read-only Long.
- `_Default As TableStyleElement  (read-only)`
- `_NewEnum As IUnknown  (read-only)`

## Methods (1)

- `Item(Index As XlTableStyleElementType) As TableStyleElement`  
  Returns a single object from a collection.
    - `Index As XlTableStyleElementType` (required): A table style element.
