# Walls

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {A2E94180-7564-4D97-806B-BBC0D0A1350C}  

Represents the walls of a 3D chart.

**Remarks:** This object is not a collection. There is no object that represents a single wall; you must return all the walls as a unit.

**Example:**

```vba
With ActiveDocument.InlineShapes(1)
 If .HasChart Then
 .Chart.Walls.Interior.Pattern = xlGray75
 End If
End With
```

## Properties (8)

- `Name As String  (read-only)`  
  Returns the name of the object. Read-only String.
- `Parent As Object  (read-only)`  
  Returns the parent for the specified object. Read-only Object.
- `PictureType As Variant  (read/write)`  
  Returns or sets a value that specifies how pictures are displayed on the walls and faces of a 3D chart. Read/write Variant.
- `PictureUnit As Variant  (read/write)`  
  Returns or sets the unit for each picture on the chart if the PictureType property is set to xlStackScale; otherwise, this property is ignored. Read/write Long.
- `Thickness As Long  (read/write)`  
  Returns or sets the thickness of the wall. Read/write Long.
- `Format As ChartFormat  (read-only)`  
  Returns the line, fill, and effect formatting for the object. Read-only ChartFormat.
- `Application As Object  (read-only)`  
  When used without an object qualifier, returns an Application object that represents the Microsoft Word application. When used with an object qualifier, returns an Application object that represents the creator of the specified object (you can use this property with an Automation object to return the application of that object). Read-only.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.

## Methods (3)

- `Select() As Variant`  
  Selects the object.
- `ClearFormats() As Variant`  
  Clears the formatting of the object.
- `Paste()`  
  Pastes a picture from the Clipboard on the walls of the specified chart.
