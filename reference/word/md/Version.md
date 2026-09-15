# Version

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {000209B4-0000-0000-C000-000000000046}  

Represents a single version of a document. The Version object is a member of the Versions collection. The Versions collection includes all the versions of the specified document.

**Remarks:** Use Versions (Index), where Index is the index number, to return a single Version object. The index number represents the position of the version in the Versions collection. The first version added to the Versions collection is index number 1. The following example displays the comment, author, and date of the first version of the active document. Use the Save method to add an item to the Versions collection. The following example adds a version of the active document with the specified comment.

## Properties (7)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Version object.
- `SavedBy As String  (read-only)`  
  Returns the name of the user who saved the specified version of the document. Read-only String.
- `Comment As String  (read-only)`  
  Returns the comment associated with the specified version of a document. Read-only String.
- `Date As Date  (read-only)`  
  The date and time that the document version was saved. Read-only Date.
- `Index As Long  (read-only)`  
  Returns a Long that represents the position of an item in a collection. Read-only.

## Methods (2)

- `Delete()`  
  Deletes the specified version.
- `Open() As Document`  
  Opens the specified version of a document. Returns a Document object representing the opened document.
