# ColorFormat

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000C0312-0000-0000-C000-000000000046}  

Represents the color of a one-color object, the foreground or background color of an object with a gradient or patterned fill, or the pointer color.

**Remarks:** You can set colors to an explicit red-green-blue value (by using the RGB property) or to a color in the color scheme (by using the SchemeColor property). Use one of the properties listed in the following table to return a ColorFormat object.

**Example:**

```vba
Set myDocument = Worksheets(1)
With myDocument.Shapes.AddShape(msoShapeRectangle, _
 90, 90, 90, 50).Fill
 .ForeColor.RGB = RGB(128, 0, 0)
 .BackColor.RGB = RGB(170, 170, 170)
 .TwoColorGradient msoGradientHorizontal, 1
End With
```

## Properties (9)

- `Application As Object  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `RGB As MsoRGBType  (read/write)`  
  Returns or sets a Long value that represents the red-green-blue value of the specified color.
- `SchemeColor As Long  (read/write)`  
  Returns or sets an Integer value that represents the color of a Color object, as an index in the current color.
- `Type As MsoColorType  (read-only)`  
  Returns an MsoColorType value that represents the color format type.
- `TintAndShade As Single  (read/write)`  
  Returns or sets a Single that lightens or darkens a color.
- `ObjectThemeColor As MsoThemeColorIndex  (read/write)`  
  Returns or sets a color that is mapped to the theme color scheme. Read/write MsoThemeColorIndex.
- `Brightness As Single  (read/write)`  
  Returns or sets the luminosity of the specified object. Read/write.
