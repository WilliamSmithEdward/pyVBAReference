# Floor

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {92D41A68-F07E-4CA4-AF6F-BEF486AA4E6F}  

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
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Application As Application  (read-only)`  
  When used without an object qualifier, returns an Application object that represents the Microsoft PowerPoint application. When used with an object qualifier, returns an Application object that represents the creator of the specified object (you can use this property with an Automation object to return the application of that object). Read-only.

## Methods (3)

- `Select() As Variant`  
  Selects the object.
- `ClearFormats() As Variant`  
  Clears the formatting of the object.
- `Paste()`  
  Pastes a picture from the Clipboard on the floor of the specified chart.
