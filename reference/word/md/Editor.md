# Editor

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {DD947D72-F33C-4198-9BDF-F86181D05E41}  

Represents a single user who has been given specific permissions to edit portions of a document.

**Remarks:** Users who can be given permissions include individual contributors and groups of users as defined for Document Workspace sites. The permissions you assign to ranges and selections go into effect only after a document is protected. Use the Editors collection and the Editor object to assign specific permissions to sections of a document. Then use the Protect method to protect the document. Use the Add method of the Editors collection to give a specified user or group permission to modify a range or selection within a document. The following example gives the current user editing permission to modify the active selection.

## Properties (7)

- `ID As String  (read-only)`  
  Returns or sets the identifying label for the specified object when the parent document is saved as a webpage. Read-only String.
- `Name As String  (read-only)`  
  Returns or sets the name of the specified object. Read-only String.
- `Range As Range  (read-only)`  
  Returns a Range object that represents the portion of a document that is contained in the specified object.
- `NextRange As Range  (read-only)`  
  Returns a Range object that represents the next range for which a user has permissions to modify.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Editor object.

## Methods (3)

- `Delete()`  
  Deletes the specified Editor object.
- `DeleteAll()`  
  Deletes all editing permissions in a document for a specific user.
- `SelectAll()`  
  Selects all the shapes in a document that were inserted or edited by a single user.
