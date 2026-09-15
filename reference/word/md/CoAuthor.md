# CoAuthor

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {E59544D5-C299-46A0-84C1-C51AB38F9759}  

Represents a single co author in the document. The CoAuthor object is a member of the CoAuthors collection. The CoAuthors collection contains all the co authors in the document (authors that are actively editing the document).

**Remarks:** Use CoAuthors (_index_), where _index_ is the index number to return a single CoAuthor object.

**Example:**

```vba
Dim author As CoAuthor

Set author = ActiveDocument.CoAuthoring.Authors(1)
MsgBox "The name of the first co author in this document is " & author.Name
```

## Properties (8)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application. Read-only.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified CoAuthor object.
- `ID As String  (read-only)`  
  Returns a String that specifies a unique identifier for the specified author. Read-only.
- `Name As String  (read-only)`  
  Returns a String that contains the display name of the specified co author. Read-only.
- `IsMe As Boolean  (read-only)`  
  Returns true if this author represents the current user. Read-only.
- `Locks As CoAuthLocks  (read-only)`  
  Returns a CoAuthLocks collection that represents the locks in the document that are associated with the specified co author. Read-only.
- `EmailAddress As String  (read-only)`  
  Returns a string that specifies the email address of the specified co author. Read-only.
