# Comments

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020940-0000-0000-C000-000000000046}  

A collection of Comment objects that represent the comments in a selection, range, or document.

**Remarks:** Use the Comments property to return the Comments collection. The following example displays comments made by Don Funk in the active document. Use the Add method to add a comment at the specified range. The following example adds a comment immediately after the selection. Use Comments (Index), where Index is the index number, to return a single Comment object. The index number represents the position of the comment in the specified selection, range, or document. The following example displays the author of the first comment in the active document. The following example displays the initials of the author of the first comment in the selection.

## Properties (6)

- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of items in the Comments collection. Read-only.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Comments object.
- `ShowBy As String  (read/write)`  
  Returns or sets the name of the reviewer whose comments are shown in the comments pane. Read/write String.

## Methods (2)

- `Item(Index As Long) As Comment`  
  Returns an individual Comment object in a collection.
    - `Index As Long` (required): The individual object to be returned. Can be a Long indicating the ordinal position of the individual object.
- `Add(Range As Range, [Text As Variant]) As Comment`  
  Returns a Comment object that represents a comment added to a range.
    - `Range As Range` (required): The range to have a comment added to it.
    - `Text As Variant` (optional): The text of the comment.
