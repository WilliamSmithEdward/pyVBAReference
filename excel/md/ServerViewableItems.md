# ServerViewableItems

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244A4-0000-0000-C000-000000000046}  

A collection of objects that have been marked as viewable on the server.

**Remarks:** This is a collection of references to objects in the workbook. Only objects in this collection will be shown on the server. By default, the entire workbook (including all worksheets) is shown. Only one ServerViewableItems object can exist per workbook. This collection is not indexable by name because there is no guarantee that the names of objects that are marked as viewable on the server are unique. In the Excel user interface, you can view the collection of objects that are marked as viewable on the server in the Excel Services Options dialog box.

## Properties (6)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Count As Long  (read-only)`  
  Returns the number of objects in the collection. Read-only Long.
- `_Default As Object  (read-only)`
- `_NewEnum As IUnknown  (read-only)`

## Methods (4)

- `Add(Obj As Variant) As Object`  
  Adds a reference to the ServerViewableItems collection.
    - `Obj As Variant` (required): The reference to an object. The object can be a reference to sheets or named items (for example, named ranges, charts, tables, and PivotTables). You cannot have both sheets and named items in the same collection.
- `Delete(Index As Variant)`  
  Deletes a reference to an object in the ServerViewableItems collection in the workbook.
    - `Index As Variant` (required): The index of the object that you want to delete.
- `DeleteAll()`  
  Deletes references to all the objects in the ServerViewableItems collection in the workbook.
- `Item(Index As Variant) As Object`  
  Returns a single object from the ServerViewableItems collection.
    - `Index As Variant` (required): The index of the object to be returned.
