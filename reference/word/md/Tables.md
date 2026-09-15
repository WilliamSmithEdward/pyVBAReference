# Tables

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {0002094D-0000-0000-C000-000000000046}  

A collection of Table objects that represent the tables in a selection, range, or document.

**Remarks:** Use the Tables property to return the Tables collection. The following example applies a border around each of the tables in the active document. Use the Add method to add a table at the specified range. The following example adds a 3x4 table at the beginning of the active document. Use Tables (Index), where Index is the index number, to return a single Table object. The index number represents the position of the table in the selection, range, or document. The following example converts the first table in the active document to text. The Count property for this collection in a document returns the number of items in the main story only. To count items in other stories use the collection with the Range object.

## Properties (6)

- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of tables in the collection. Read-only.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Tables object.
- `NestingLevel As Long  (read-only)`  
  Returns the nesting level of the specified tables. Read-only Long.

## Methods (2)

- `Item(Index As Long) As Table`  
  Returns an individual Table object in a collection.
    - `Index As Long` (required): The individual object to be returned. Can be a Long indicating the ordinal position of the individual object.
- `Add(Range As Range, NumRows As Long, NumColumns As Long, [DefaultTableBehavior As Variant], [AutoFitBehavior As Variant]) As Table`  
  Returns a Table object that represents a new, blank table added to a document.
    - `Range As Range` (required): The range where you want the table to appear. The table replaces the range, if the range isn't collapsed.
    - `NumRows As Long` (required): The number of rows you want to include in the table.
    - `NumColumns As Long` (required): The number of columns you want to include in the table.
    - `DefaultTableBehavior As Variant` (optional): Sets a value that specifies whether Microsoft Word automatically resizes cells in tables to fit the cells' contents (AutoFit). Can be either of the following constants: wdWord8TableBehavior (AutoFit disabled) or wdWord9TableBehavior (AutoFit enabled). The default constant is wdWord8TableBehavior.
    - `AutoFitBehavior As Variant` (optional): Sets the AutoFit rules for how Word sizes tables. Can be one of the WdAutoFitBehavior constants.
