# Comment

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {0002093D-0000-0000-C000-000000000046}  

Represents a single comment. The Comment object is a member of the Comments collection. The Comments collection includes comments in a selection, range or document.

**Remarks:** Use Comments (Index), where Index is the index number, to return a single Comment object. The index number represents the position of the comment in the specified selection, range, or document. The following example displays the author of the first comment in the active document. Use the Add method to add a comment at the specified range. The following example adds a comment immediately after the selection. Use the Reference property to return the reference mark associated with the specified comment. Use the Range property to return the text associated with the specified comment. The following example displays the text associated with the first comment in the active document.

## Properties (13)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Comment object.
- `Range As Range  (read-only)`  
  Returns a Range object that represents the contents of a comment.
- `Reference As Range  (read-only)`  
  Returns a Range object that represents a reference mark for a comment.
- `Scope As Range  (read-only)`  
  Returns a Range object that represents the range of text marked by the specified comment.
- `Index As Long  (read-only)`  
  Returns a Long that represents the position of an item in a collection. Read-only.
- `Date As Date  (read-only)`  
  Returns a Date that represents the date and time that a comment was inserted. Read-only.
- `IsInk As Boolean  (read-only)`  
  Returns a Boolean that represents whether a comment is a handwritten comment.
- `Done As Boolean  (read/write)`  
  Returns or sets a Boolean whose value is True if the specified comment has been marked closed. Read/write.
- `Ancestor As Comment  (read-only)`  
  For comments that are replies to existing comments, returns the parent Comment object; for new (top-level) comments, returns Nothing. Read-only.
- `Contact As CoAuthor  (read-only)`  
  Returns a CoAuthor object that represents the author of the specified comment. Read-only.
- `Replies As Comments  (read-only)`  
  Returns a Comments collection of Comment objects that are children of the specified comment. Read-only.

## Methods (2)

- `Edit()`  
  Opens the specified OLE object for editing in the application it was created in.
- `DeleteRecursively()`  
  Deletes the specified comment and all replies associated with it.
