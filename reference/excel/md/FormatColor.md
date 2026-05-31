# FormatColor

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024491-0000-0000-C000-000000000046}  

Represents the fill color specified for a threshold of a color scale conditional format or the color of the bar in a data bar conditional format.

**Remarks:** You can choose a color by passing an RGB value in the Color property, or designate the color by indexing into the theme color palette by using the ThemeColor property.

**Example:**

```vba
Sub CreateColorScaleCF()

 Dim cfColorScale As ColorScale

 'Fill cells with sample data from 1 to 10
 With ActiveSheet
 .Range("C1") = 1
 .Range("C2") = 2
 .Range("C1:C2").AutoFill Destination:=Range("C1:C10")
 End With

 Range("C1:C10").Select

 'Create a two-color ColorScale object for the created sample data range
 Set cfColorScale = Selection.FormatConditions.AddColorScale(ColorScaleType:=2)

 'Set the minimum threshold to red and maximum threshold to blue
 cfColorScale.ColorScaleCriteria(1).FormatColor.Color = RGB(255, 0, 0)
 cfColorScale.ColorScaleCriteria(2).FormatColor.Color = RGB(0, 0, 255)

End Sub
```

## Properties (7)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Color As Variant  (read/write)`  
  Returns or sets the fill color associated with a threshold for a data bar or color scale conditional formatting rule.
- `ColorIndex As XlColorIndex  (read/write)`  
  Returns or sets one of the constants of the XlColorIndex enumeration, specifying if the fill color is expressed as an index value into the current color palette.
- `ThemeColor As Variant  (read/write)`  
  Returns or sets one of the constants of the XlThemeColor enumeration, specifying the theme color used in a threshold of a data bar or color scale conditional format.
- `TintAndShade As Variant  (read/write)`  
  Returns or sets a Single that lightens or darkens the fill color of a cell for a threshold of a data bar or color scale conditional formatting rule.
