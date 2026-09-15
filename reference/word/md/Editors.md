# Editors

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {AED7E08C-14F0-4F33-921D-4C5353137BF6}  

A collection of Editor objects that represents a collection of users or groups of users who have been given specific permissions to edit portions of a document.

**Remarks:** Use the Add method to give a specified user or group permission to modify a range or selection within a document. The following example gives the current user editing permission to modify the active selection.

## Properties (4)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Editors object.
- `Count As Long  (read-only)`  
  Returns a the Long that represents the number of Editor objects in the collection. Read-only.

## Methods (2)

- `Item(Index As Variant) As Editor`  
  Returns an Editor object that represents a specific user or a group of users who have been given permission to edit a portion of a document.
    - `Index As Variant` (required): Can be either a String that represents the user's email alias (if in the same domain), an email address, or a WdEditorType constant that represents a group of users.
- `Add(EditorID As Variant) As Editor`  
  Returns an Editor object that represents a new permission for a specified user to modify a range or selection within a document. .
    - `EditorID As Variant` (required): Can be either a String that represents the user's email alias (if in the same domain), an email address, or a WdEditorType that represents a group of users.
