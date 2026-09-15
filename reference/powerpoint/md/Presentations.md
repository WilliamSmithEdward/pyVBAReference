# Presentations

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {91493462-5A91-11CF-8700-00AA0060263B}  

A collection of all the Presentation objects in Microsoft PowerPoint. Each Presentation object represents a presentation that's currently open in PowerPoint.

**Remarks:** The Presentations collection doesn't include open add-ins, which are a special kind of hidden presentation. You can, however, return a single open add-in if you know its file name. For example Presentations("oscar.ppa") will return the open add-in named "Oscar.ppa" as a Presentation object. However, it is recommended that the AddIns collection be used to return open add-ins. If your Visual Studio solution includes the Microsoft.Office.Interop.PowerPoint reference, this collection maps to the following types: - Microsoft.Office.Interop.PowerPoint.Presentations.GetEnumerator (to enumerate the Presentation objects.)

**Example:**

```vba
Set newPres = Presentations.Add(True)
newPres.Slides.Add 1, 1
newPres.SaveAs "Sample"
```

## Properties (3)

- `Count As Long  (read-only)`  
  Returns the number of objects in the specified collection. Read-only.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.

## Methods (6)

- `Item(Index As Variant) As Presentation`  
  Returns a single Presentation object from the specified Presentations collection.
    - `Index As Variant` (required): The name or index number of the single Presentation object in the collection to be returned.
- `Add([WithWindow As MsoTriState]) As Presentation`  
  Creates a presentation. Returns a Presentation object that represents the new presentation.
    - `WithWindow As MsoTriState` (optional): Whether the presentation appears in a visible window.
- `Open(FileName As String, [ReadOnly As MsoTriState], [Untitled As MsoTriState], [WithWindow As MsoTriState]) As Presentation`  
  Opens the specified presentation. Returns a Presentation object that represents the opened presentation.
    - `FileName As String` (required): The name of the file to open.
    - `ReadOnly As MsoTriState` (optional): Specifies whether the file is opened with read/write or read-only status.
    - `Untitled As MsoTriState` (optional): Specifies whether the file has a title.
    - `WithWindow As MsoTriState` (optional): Specifies whether the file is visible.
- `CheckOut(FileName As String)`  
  Copies a specified presentation from a server to a local computer for editing. Returns a String that represents the local path and file name of the presentation checked out.
    - `FileName As String` (required): The server path and name of the presentation.
- `CanCheckOut(FileName As String) As Boolean`  
  Returns True if Microsoft PowerPoint can check out a specified presentation from a server.
    - `FileName As String` (required): The server path and name of the presentation.
- `Open2007(FileName As String, [ReadOnly As MsoTriState], [Untitled As MsoTriState], [WithWindow As MsoTriState], [OpenAndRepair As MsoTriState]) As Presentation`  
  Opens the specified presentation and provides the option to repair the presentation file. Returns a Presentation object that represents the opened presentation.
    - `FileName As String` (required): The name of the file to open.
    - `ReadOnly As MsoTriState` (optional): Specifies whether the file is opened with read/write or read-only status.
    - `Untitled As MsoTriState` (optional): Specifies whether the file has a title.
    - `WithWindow As MsoTriState` (optional): Specifies whether the file is visible.
    - `OpenAndRepair As MsoTriState` (optional): Specifies whether to repair the file before it is opened to prevent corruption.
