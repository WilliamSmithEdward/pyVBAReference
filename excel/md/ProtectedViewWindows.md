# ProtectedViewWindows

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244CC-0000-0000-C000-000000000046}  

A collection of the ProtectedViewWindow objects that represent all the Protected View windows that are currently open in the application.

**Remarks:** Use the ProtectedViewWindows property of the Application object to return the ProtectedViewWindows collection.

**Example:**

```vba
MsgBox "There are " & ProtectedViewWindows.Count & _
 " Protected View windows currently open."
```

## Properties (7)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Count As Long  (read-only)`  
  Returns a Long value that represents the number of objects in the collection.
- `Item As ProtectedViewWindow  (read-only)`  
  Returns a single object from a collection.
- `_NewEnum As IUnknown  (read-only)`
- `_Default As ProtectedViewWindow  (read-only)`

## Methods (1)

- `Open(Filename As String, [Password As Variant], [AddToMru As Variant], [RepairMode As Variant]) As ProtectedViewWindow`  
  Opens the specified workbook in a new Protected View window.
    - `Filename As String` (required): The name of the workbook (paths are accepted).
    - `Password As Variant` (optional): The password for opening the workbook.
    - `AddToMru As Variant` (optional): True to add the file name to the list of recently used files on the Recent tab of the Backstage view.
    - `RepairMode As Variant` (optional): True to repair the workbook to prevent file corruption.
