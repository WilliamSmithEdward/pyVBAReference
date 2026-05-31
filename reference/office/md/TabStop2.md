# TabStop2

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C03BB-0000-0000-C000-000000000046}  

Represents a single tab stop. The TabStop2 object is a member of the TabStops2 collection.

**Remarks:** Tab stops are indexed numerically from left to right along the ruler.

**Example:**

```vba
Sub ClearTabStop()
 Selection.TextRange.ParagraphFormat.Tabs(1).Clear
End Sub
```

## Properties (5)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the TabStop2 object. Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that represents the Microsoft Office application in which the TabStop2 object was created. Read-only.
- `Parent As Object  (read-only)`  
  Gets an Object that represents the Parent object for the TabStop2 object. Read-only.
- `Position As Single  (read/write)`  
  Gets or sets the position of a tab stop relative to the left margin. Read/write.
- `Type As MsoTabStopType  (read/write)`  
  Gets or sets the type of the TabStop2 object. Read/write.

## Methods (1)

- `Clear()`  
  Removes the specified custom tab stop.
