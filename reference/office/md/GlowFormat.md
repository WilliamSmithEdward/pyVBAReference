# GlowFormat

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C03BD-0000-0000-C000-000000000046}  

Represents a glow effect around an Office graphic.

**Example:**

```vba
With ActivePresentation.Slides(2).Shapes(2)
 .Text.Font.Glowformat = msoGlowType2
End With
```

## Properties (5)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the GlowFormat object. Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the GlowFormat object was created. Read-only.
- `Radius As Single  (read/write)`  
  Gets or sets the radius value of the glow effect for the GlowFormat object. Read/write.
- `Color As ColorFormat  (read-only)`  
  Gets a ColorFormat object that represents the color of text formatted as glow. Read-only.
- `Transparency As Single  (read/write)`  
  Returns or sets the degree of transparency of the specified glow as a value between 0.0 (opaque) and 1.0 (clear). Read/write.
