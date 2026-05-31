# TabStops2

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C03BA-0000-0000-C000-000000000046}  

The collection of TabStop2 objects.

**Remarks:** Tab stops are indexed numerically from left to right along the ruler.

**Example:**

```vba
Sub ClearTabStop()
    ActiveDocument.Pages(1).Shapes(1).TextFrame.TextRange _
        .ParagraphFormat.Tabs(1).Clear
End Sub
```

## Properties (6)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the TabStops2 object. Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that represents the Microsoft Office application in which the TabStops2 object was created. Read-only.
- `Parent As Object  (read-only)`  
  Gets an Object that represents the Parent object of the specified TabStops2 object. Read-only.
- `Count As Long  (read-only)`  
  Gets a Long indicating the number of items in the TabStops2 collection. Read-only.
- `DefaultSpacing As Single  (read/write)`  
  Gets or sets the default spacing between tab stops. Read/write.
- `_NewEnum As IUnknown  (read-only)`

## Methods (2)

- `Item(Index As Variant) As TabStop2`  
  Gets an individual object from the TabStops2 collection.
    - `Index As Variant` (required): The number of the object to return.
- `Add(Type As MsoTabStopType, Position As Single) As TabStop2`  
  Adds a new tab stop to the specified TabStops2 object.
    - `Type As MsoTabStopType` (required): The type of tab stop to add.
    - `Position As Single` (required): The horizontal position of the new tab stop relative to the left edge of the text frame. Numeric values are evaluated in points; strings are evaluated in the units specified and can be in any measurement unit supported by the Microsoft Office product.
