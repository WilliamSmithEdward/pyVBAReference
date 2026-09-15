# Subdocument

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020989-0000-0000-C000-000000000046}  

Represents a subdocument within a document or range. The Subdocument object is a member of the Subdocuments collection. The Subdocuments collection includes all the subdocuments in the a range or document.

**Remarks:** Use Subdocuments (Index), where Index is the index number, to return a single Subdocument object. The following example displays the path and file name of the first subdocument in the active document. Use the AddFromFile or AddFromRange method to add a subdocument to a document. The following example adds a subdocument named "Setup.doc" at the end of the active document. The following example applies the Heading 1 style to the first paragraph in the selection and then creates a subdocument for the contents of the selection.

## Properties (9)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Subdocument object.
- `Locked As Boolean  (read/write)`  
  True if a subdocument in a master document is locked. Read/write Boolean.
- `Range As Range  (read-only)`  
  Returns a Range object that represents the portion of a document that's contained within the subdocument.
- `Name As String  (read-only)`  
  Returns the name of the specified object. Read-only String.
- `Path As String  (read-only)`  
  Returns the disk or Web path to the specified subdocument. Read-only String.
- `HasFile As Boolean  (read-only)`  
  True if the specified subdocument has been saved to a file. Read-only Boolean.
- `Level As Long  (read-only)`  
  Returns the heading level used to create the subdocument. Read-only Long.

## Methods (3)

- `Delete()`  
  Deletes the specified subdocument.
- `Split(Range As Range)`  
  Divides an existing subdocument into two subdocuments at the same level in master document view or outline view.
    - `Range As Range` (required): The range that, when the subdocument is split, becomes a separate subdocument.
- `Open() As Document`  
  Opens the specified subdocument. Returns a Document object that represents the opened subdocument.
