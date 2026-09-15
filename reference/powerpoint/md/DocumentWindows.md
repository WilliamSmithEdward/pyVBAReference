# DocumentWindows

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {91493455-5A91-11CF-8700-00AA0060263B}  

A collection of all the DocumentWindow objects that are currently open in Microsoft PowerPoint. This collection doesn't include open slide show windows, which are included in the SlideShowWindows collection.

**Remarks:** If your Visual Studio solution includes the Microsoft.Office.Interop.PowerPoint reference, this collection maps to the following types: - Microsoft.Office.Interop.PowerPoint.DocumentWindows

**Example:**

```vba
Windows.Arrange ppArrangeTiled
```

## Properties (3)

- `Count As Long  (read-only)`  
  Returns the number of objects in the specified collection. Read-only.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.

## Methods (2)

- `Item(Index As Long) As DocumentWindow`  
  Returns a single DocumentWindow object from the specified DocumentWindows collection.
    - `Index As Long` (required): The index number of the single DocumentWindow object in the collection to be returned.
- `Arrange([arrangeStyle As PpArrangeStyle])`  
  Arranges all open document windows in the workspace.
    - `arrangeStyle As PpArrangeStyle` (optional): Specifies whether to cascade or tile the windows.
