# ColorStop

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244AD-0000-0000-C000-000000000046}  

Represents the color stop point for a gradient fill in a range or selection.

**Remarks:** The ColorStop object enables you to set properties for the cell fill, including the Color, ThemeColor, and TintAndShade properties.

**Example:**

```vba
With Selection.Interior
 .Pattern = xlPatternLinearGradient
 .Gradient.Degree = 135
 .Gradient.ColorStops.Clear
End With

With Selection.Interior.Gradient.ColorStops.Add(0)
 .ThemeColor = xlThemeColorDark1
 .TintAndShade = 0
End With

With Selection.Interior.Gradient.ColorStops.Add(0.5)
 .ThemeColor = xlThemeColorAccent1
 .TintAndShade = 0
End With

With Selection.Interior.Gradient.ColorStops.Add(1)
 .ThemeColor = xlThemeColorDark1
 .TintAndShade = 0
End With
```

## Properties (7)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Color As Variant  (read/write)`  
  Returns or sets the color of the represented object. Read/write.
- `ThemeColor As Long  (read/write)`  
  Returns or sets the theme color of the represented object. Read/write.
- `TintAndShade As Variant  (read/write)`  
  Returns or sets the tint and shade of the represented object. Read/write
- `Position As Double  (read/write)`  
  Returns or sets the position of the ColorStop. Read/write.

## Methods (1)

- `Delete()`  
  Deletes the represented object.
