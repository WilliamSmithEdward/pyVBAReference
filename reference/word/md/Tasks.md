# Tasks

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020983-0000-0000-C000-000000000046}  

A collection of Task objects that represents all the tasks currently running on the system.

**Remarks:** Use the Tasks property to return the Tasks collection. The following example determines whether Microsoft Excel is running. If it is, this example switches to it and maximizes it; otherwise, the example starts it. Use Visual Basic's Shell function to run an executable program and add the program to the Tasks collection. Use Tasks (Index), where Index is the application name or the index number, to return a single Task object. The following example opens and resizes the application window for the first visible task in the Tasks collection. The following example restores the Calculator application window if the application is in the Tasks collection.

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Tasks object.
- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of tasks in the collection. Read-only.

## Methods (3)

- `Item(Index As Variant) As Task`  
  Returns an individual Task object in a collection.
    - `Index As Variant` (required): The individual object to be returned. Can be a Long indicating the ordinal position or a String representing the name of the individual object.
- `Exists(Name As String) As Boolean`  
  Determines whether the specified task exists. Returns True if the task exists.
    - `Name As String` (required): The name of the task.
- `ExitWindows()`  
  Closes all open applications, quits Microsoft Windows, and logs the current user off.
