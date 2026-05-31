# ReflectionFormat

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C03BE-0000-0000-C000-000000000046}  

Represents the reflection effect in Office graphics.

**Example:**

```vba
With ActivePresentation.Slides(1).Shapes(2)
 With .TextFrame2.TextRange.Font
 .Size = 32
 .Name = "Palatino"
 .Reflection.Type = msoReflectionType6
 End With
End With
```

## Properties (7)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the ReflectionFormat object. Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the ReflectionFormat object was created. Read-only.
- `Type As MsoReflectionType  (read/write)`  
  Gets or sets the type of the ReflectionFormat object. Read/write.
- `Transparency As Single  (read/write)`  
  Gets or sets the amount of transparency, measured in percentages, of the shape's reflection image. Read/write.
- `Size As Single  (read/write)`  
  Gets or sets the size, measured in percentages, of the shape's reflection image. Read/write.
- `Offset As Single  (read/write)`  
  Gets or sets the amount of separation, measured in points, of the reflected image from the shape. Read/write.
- `Blur As Single  (read/write)`  
  Gets or sets the amount of blur, measured in points, of the shape's reflection image. Read/write.
