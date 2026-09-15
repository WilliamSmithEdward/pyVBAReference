# ProtectedViewWindows

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {FD0A74E8-C719-49F6-BA1B-F6D9839D1AB9}  

A collection of all the ProtectedViewWindow objects that are currently open in Word.

**Remarks:** Use the ProtectedViewWindows property to return the ProtectedViewWindows collection.

**Example:**

```vba
MsgBox "There are " & ProtectedViewWindows.Count & _
 " Protected View windows currently open."
```

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application. Read-only.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified ProtectedViewWindows object.
- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of Protected View windows in the collection. Read-only.

## Methods (2)

- `Item(Index As Variant) As ProtectedViewWindow`  
  Returns an individual ProtectedViewWindows object in a collection.
    - `Index As Variant` (required): The individual object to be returned.
- `Open(FileName As Variant, [AddToRecentFiles As Variant], [PasswordDocument As Variant], [Visible As Variant], [OpenAndRepair As Variant]) As ProtectedViewWindow`  
  Opens the specified document in a new Protected View window.
    - `FileName As Variant` (required): The name of the document (paths are accepted).
    - `AddToRecentFiles As Variant` (optional): True to add the file name to the list of recently used files at the bottom of the File menu.
    - `PasswordDocument As Variant` (optional): The password for opening the document.
    - `Visible As Variant` (optional): True if the document is opened in a visible window. The default value is True.
    - `OpenAndRepair As Variant` (optional): True to repair the document to prevent document corruption.
