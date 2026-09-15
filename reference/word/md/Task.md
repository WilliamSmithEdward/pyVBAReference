# Task

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020982-0000-0000-C000-000000000046}  

Represents a single task running on the system. The Task object is a member of the Tasks collection. The Tasks collection includes all the applications that are currently running on the system.

**Remarks:** Use Tasks (Index), where Index is the application name or the index number, to return a single Task object. The following example switches to and resizes the application window for the first visible task in the Tasks collection. The following example restores the Calculator application window if Calculator is in the Tasks collection. Use Visual Basic's Shell function to run an executable program and add the program to the Tasks collection.

## Properties (10)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Task object.
- `Name As String  (read-only)`  
  Returns the name of the specified object. Read-only String.
- `Left As Long  (read/write)`  
  Returns or sets a Long that represents the horizontal position of the specified task, measured in points. Read/write.
- `Top As Long  (read/write)`  
  Returns or sets the vertical position, in points, of the specified window. Read/write Long.
- `Width As Long  (read/write)`  
  Returns or sets the width of the specified Task object, in points. Read/write Long.
- `Height As Long  (read/write)`  
  Returns or sets the height of the specified task window. Read/write Long.
- `WindowState As WdWindowState  (read/write)`  
  Returns or sets the state of the specified document window or task window. Read/write WdWindowState.
- `Visible As Boolean  (read/write)`  
  True if the specified object is visible. Read/write Boolean.

## Methods (5)

- `Activate([Wait As Variant])`  
  Activates the Task object.
    - `Wait As Variant` (optional): True to wait until the user has activated Word before activating the task. False to immediately activate the task, even if Word isn't active.
- `Close()`  
  Closes the specified task.
- `Move(Left As Long, Top As Long)`  
  Positions a task window.
    - `Left As Long` (required): The horizontal screen position of the specified window.
    - `Top As Long` (required): The vertical screen position of the specified window.
- `Resize(Width As Long, Height As Long)`  
  Sizes the specified task window.
    - `Width As Long` (required): The width of the window, in points.
    - `Height As Long` (required): The height of the window, in points.
- `SendWindowMessage(Message As Long, wParam As Long, lParam As Long)`  
  Sends a Windows message and its associated parameters to the specified task.
    - `Message As Long` (required): A hexadecimal number that corresponds to the message you want to send. If you have the Microsoft Platform Software Development Kit, you can look up the name of the message in the header files (Winuser.h, for example) to find the associated hexadecimal number (precede the hexadecimal value with &h).
    - `wParam As Long` (required): Parameters appropriate for the message you are sending. For information about what these values represent, see the reference topic for that message in the documentation included with the Microsoft Platform Software Development Kit, available on MSDN. To retrieve the appropriate values, you may need to use the Spy tool (which comes with the kit).
