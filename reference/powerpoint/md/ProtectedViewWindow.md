# ProtectedViewWindow

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {BA72E55A-4FF5-48F4-8215-5505F990966F}  

Represents a presentation window that is in protected view.

**Remarks:** Documents displayed in a Protected View window cannot be edited and are restricted from running active content such as Visual Basic for Applications macros and data connections. Use ProtectedViewWindows (_index_), where _index_ is the index number to return a single ProtectedViewWindow object. The index number represents the position of the Protected View window in the ProtectedViewWindows collection.

## Properties (12)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Presentation As Presentation  (read-only)`  
  Returns a Presentation object that represents the presentation in which the specified window was created. Read-only.
- `Active As MsoTriState  (read-only)`  
  Returns whether the specified window is active. Read-only.
- `WindowState As PpWindowState  (read/write)`  
  Returns or sets the state of the specified window. Read/write.
- `Caption As String  (read-only)`  
  Returns the text that appears in the title bar of the specified window. Read-only.
- `SourcePath As String  (read-only)`  
  Returns the source path of the current ProtectedViewWindow object. Read-only.
- `SourceName As String  (read-only)`  
  Returns the source name of the current ProtectedViewWindow object. Read-only.
- `Left As Single  (read/write)`  
  Returns or sets the distance in points from the left edge of the specified window to the left edge of the application window's client area. Read/write.
- `Top As Single  (read/write)`  
  Returns or sets the distance in points from the top edge of the specified window to the top edge of the application window's client area. Read/write.
- `Width As Single  (read/write)`  
  Returns or sets the width of the specified object, in points. Read/write.
- `Height As Single  (read/write)`  
  Returns or sets the height of the specified object, in points. Read/write.

## Methods (3)

- `Activate()`  
  Activates the specified window.
- `Close()`  
  Closes the specified window.
- `Edit([ModifyPassword As String]) As Presentation`  
  Provides the password of the ProtectedViewWindow object to make editing possible.
    - `ModifyPassword As String` (optional): The modification password as set in the file.
