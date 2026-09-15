# ProtectedViewWindow

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {F743EDD0-9B97-4B09-89CC-77BE19B51481}  

Represents a Protected View window.

**Remarks:** Documents displayed in a Protected View window cannot be edited and are restricted from running active content such as Visual Basic for Applications macros and Data Connections. Use ProtectedViewWindows (_index_), where _index_ is the index number to return a single ProtectedViewWindow object.

**Example:**

```vba
Dim pvWindow As ProtectedViewWindow

Set pvWindow = ProtectedViewWindows(1)
```

## Properties (15)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application. Read-only.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified ProtectedViewWindow object.
- `Caption As String  (read/write)`  
  Returns or sets the caption text that is displayed in the title bar of the document or Protected View window. Read/write String.
- `Document As Document  (read-only)`  
  Returns a Document object associated with the Protected View window. Read-only.
- `Left As Long  (read/write)`  
  Returns or sets a Long, in points, that represents the horizontal position of the specified Protected View window. Read/write.
- `Top As Long  (read/write)`  
  Returns or sets the vertical position, in points, of the specified Protected View window. Read/write Long
- `Width As Long  (read/write)`  
  Returns or sets the width, in points, of the specified Protected View window. Read/write Long.
- `Height As Long  (read/write)`  
  Returns or sets the height of the Protected View window. Read/write Long.
- `WindowState As WdWindowState  (read/write)`  
  Returns or sets the state of the specified Protected View window. Read/write WdWindowState.
- `Active As Boolean  (read-only)`  
  True if the specified Protected View window is active. Read-only Boolean.
- `Index As Long  (read-only)`  
  Returns a Long that represents the position of an item in a collection. Read-only.
- `Visible As Boolean  (read/write)`  
  True if the specified Protected View window is visible. Read/write. Boolean.
- `SourceName As String  (read-only)`  
  Returns the name of the source file for the specified Protected View window. Read-only String.
- `SourcePath As String  (read-only)`  
  Returns the path of the source file for the specified Protected View window. Read-only String.

## Methods (4)

- `Activate()`  
  Activates the specified Protected View window.
- `Edit([PasswordTemplate As Variant], [WritePasswordDocument As Variant], [WritePasswordTemplate As Variant]) As Document`
    - `PasswordTemplate As Variant` (optional): The password for opening the template.
    - `WritePasswordDocument As Variant` (optional): The password for saving changes to the document.
    - `WritePasswordTemplate As Variant` (optional): The password for saving changes to the template.
- `Close()`  
  Closes the specified Protected View window.
- `ToggleRibbon()`  
  Shows or hides the ribbon.
