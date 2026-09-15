# ChartBorder

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {92D41A56-F07E-4CA4-AF6F-BEF486AA4E6F}  

Represents the border of an object.

**Remarks:** Most bordered objects have a border that is treated as a single entity, regardless of how many sides it has. The entire border must be returned as a unit. To return a Border object, use the Border property for the particular bordered object (for example, the Border property of a TrendLine object).

**Example:**

```vba
With ActiveDocument.InlineShapes(1).Chart.SeriesCollection(1).Trendlines(1)

    .Type = xlLinear

    .Border.LineStyle = xlDash

End With
```

## Properties (7)

- `Color As Variant  (read/write)`  
  Returns or sets the primary color of the object. Read/write Variant.
- `ColorIndex As Variant  (read/write)`  
  Returns or sets the color of the border. Read/write Variant.
- `LineStyle As Variant  (read/write)`  
  Returns or sets the line style for the border. Read/write XlLineStyle, xlGray25, xlGray50, xlGray75, or xlAutomatic.
- `Weight As Variant  (read/write)`  
  Returns or sets the weight of the border. Read/write XlBorderWeight.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent for the specified object. Read-only Object.
- `Application As Application  (read-only)`  
  When used without an object qualifier, returns an Application object that represents the Microsoft PowerPoint application. When used with an object qualifier, returns an Application object that represents the creator of the specified object (you can use this property with an Automation object to return the application of that object). Read-only.
