# TextColumns

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020973-0000-0000-C000-000000000046}  

## Properties (10)

- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of text columns in the collection. Read-only.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified TextColumns object.
- `EvenlySpaced As Long  (read/write)`  
  True if text columns are evenly spaced. Read/write Long.
- `LineBetween As Long  (read/write)`  
  True if vertical lines appear between all the columns in the TextColumns collection. Read/write Long.
- `Width As Single  (read/write)`  
  Returns or sets the width of the Word art text effects, in points. Read/write Long.
- `Spacing As Single  (read/write)`  
  Returns or sets the spacing (in points) between columns. Read/write Single.
- `FlowDirection As WdFlowDirection  (read/write)`  
  Returns or sets the direction in which text flows from one text column to the next. Read/write WdFlowDirection.

## Methods (3)

- `Item(Index As Long) As TextColumn`  
  Returns an individual TextColumn object in a collection.
    - `Index As Long` (required): The individual object to be returned. Can be a Long indicating the ordinal position of the individual object.
- `Add([Width As Variant], [Spacing As Variant], [EvenlySpaced As Variant]) As TextColumn`  
  Returns a TextColumn object that represents a new text column added to a section or document.
    - `Width As Variant` (optional): The width of the new text column in the document, in points.
    - `Spacing As Variant` (optional): The spacing between the text columns in the document, in points.
    - `EvenlySpaced As Variant` (optional): True to evenly space all the text columns be in the document.
- `SetCount(NumColumns As Long)`  
  Arranges text into the specified number of text columns.
    - `NumColumns As Long` (required): The number of columns the text is to be arranged into.
