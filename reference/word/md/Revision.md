# Revision

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020981-0000-0000-C000-000000000046}  

Represents a change marked with a revision mark. The Revision object is a member of the Revisions collection. The Revisions collection includes all the revision marks in a range or document.

**Remarks:** Use Revisions (Index), where Index is the index number, to return a single Revision object. The index number represents the position of the revision in the range or document. The following example displays the author name for the first revision in section one of the active document. The Add method isn't available for the Revisions collection. Revision objects are added when change tracking is enabled. Set the TrackRevisions property to True to track revisions made to the document text. The following example enables revision tracking and then inserts "Action " before the selection.

## Properties (12)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Revision object.
- `Author As String  (read-only)`  
  Returns the name of the user who made the specified tracked change. Read-only String.
- `Date As Date  (read-only)`  
  The date and time that the tracked change was made. Read-only Date.
- `Range As Range  (read-only)`  
  Returns a Range object that represents the portion of a document that's contained within a revision mark.
- `Type As WdRevisionType  (read-only)`  
  Returns the revision type. Read-only WdRevisionType.
- `Index As Long  (read-only)`  
  Returns a Long that represents the position of an item in a collection. Read-only.
- `Style As Style  (read-only)`  
  Returns a Style object that represents the style associated with a revision mark.
- `FormatDescription As String  (read-only)`  
  Returns a String representing a description of tracked formatting changes in a revision. Read-only.
- `MovedRange As Range  (read-only)`  
  Returns a Range object that represents a range of text that was moved from one place to another in a document with tracked changes. Read-only.
- `Cells As Cells  (read-only)`  
  Returns a Cells collection that represents the table cells that have been marked with revision marks. Read-only.

## Methods (2)

- `Accept()`  
  Accepts the specified tracked change, removes the revision mark, and incorporates the change into the document.
- `Reject()`  
  Rejects the specified tracked change. The revision marks are removed, leaving the original text intact.
