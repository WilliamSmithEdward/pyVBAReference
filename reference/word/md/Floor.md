# Floor

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {7E64D2BE-2818-48CB-8F8A-CC7B61D9E860}  

Represents the floor of a 3D chart.

**Example:**

```vba
With ActiveDocument.InlineShapes(1)
 If .HasChart Then
 .Chart.Floor.Interior.Color = RGB(0, 255, 255)
 End If
End With
```

## Properties (7)

- `Name As String  (read-only)`  
  Returns the name of the object. Read-only String.
- `Parent As Object  (read-only)`  
  Returns the parent for the specified object. Read-only Object.
- `PictureType As Variant  (read/write)`  
  Returns or sets a value that specifies how pictures are displayed on the walls and faces of a 3D chart. Read/write Variant.
- `Thickness As Long  (read/write)`  
  Returns or sets the thickness of the floor. Read/write Long.
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
  Pastes a picture from the Clipboard on the floor of the specified chart.
