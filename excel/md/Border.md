# Border

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00020854-0000-0000-C000-000000000046}  

Represents the border of an object.

**Remarks:** Most bordered objects (all except for the Range and Style objects) have a border that's treated as a single entity, regardless of how many sides it has. The entire border must be returned as a unit. Use the Border property, such as from a TrendLine object, to return the Border object for this kind of object. Note that the visual properties of a Border object are interlocked; that is, changing one property can induce changes in another. In most cases, the induced changes serve to make the border visible (which may or may not be desirable). However, other (more unexpected) results are possible. Following is an example of interlocking with unexpected results. In this example, setting a border's Weight property to xlThick induces the LineStyle property to become xlSolid despite having previously set it to xlDashDotDot.

**Example:**

```vba
With ActiveChart.SeriesCollection(1).Trendlines(1)
 .Type = xlLinear
 .Border.LineStyle = xlDash
End With
```

## Properties (9)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Color As Variant  (read/write)`  
  Returns or sets the primary color of the object, as shown in the table in the remarks section. Use the RGB function to create a color value. Read/write Variant.
- `ColorIndex As Variant  (read/write)`  
  Returns or sets a Variant value that represents the color of the border.
- `LineStyle As Variant  (read/write)`  
  Returns or sets the line style for the border. Read/write XlLineStyle, xlGray25, xlGray50, xlGray75, or xlAutomatic.
- `Weight As Variant  (read/write)`  
  Returns or sets an XlBorderWeight value that represents the weight of the border.
- `ThemeColor As Variant  (read/write)`  
  Returns or sets the theme color in the applied color scheme that is associated with the specified object. Read/write Variant.
- `TintAndShade As Variant  (read/write)`  
  Returns or sets a Single that lightens or darkens a color.
