# TaskPane

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {B9F1A4E2-0D0A-43B7-8495-139E7ACBD840}  

Represents a single task pane available to Microsoft Word, which contains common tasks that users perform. The TaskPane object is a member of the TaskPanes collection.

**Remarks:** Use the TaskPanes property to return a TaskPane object. Use the Visible property to display an individual task pane. This example displays the formatting task pane.

## Properties (4)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified TaskPane object.
- `Visible As Boolean  (read/write)`  
  True if the specified object is visible. Read/write Boolean.
