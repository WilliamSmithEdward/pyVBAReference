# LegendKey

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {DF076FDE-8781-4051-A5BC-99F6B7DC04D4}  

Represents a legend key in a chart legend.

**Remarks:** Each legend key is a graphic that visually links a legend entry with its associated series or trendline in the chart. The legend key is linked to its associated series or trendline in such a way that changing the formatting of one simultaneously changes the formatting of the other.

**Example:**

```vba
With ActiveDocument.InlineShapes(1)
 If .HasChart Then
 .Chart.Legend.LegendEntries(1).LegendKey _
 .MarkerBackgroundColorIndex = 5
 End If
End With
```

## Properties (19)

- `Parent As Object  (read-only)`  
  Returns the parent for the specified object. Read-only Object.
- `InvertIfNegative As Boolean  (read/write)`  
  True if Microsoft Word inverts the pattern in the object when it corresponds to a negative number. Read/write Variant.
- `MarkerBackgroundColor As Long  (read/write)`  
  Sets the marker background color as an RGB value or returns the corresponding color index value. Read/write Long.
- `MarkerBackgroundColorIndex As XlColorIndex  (read/write)`  
  Returns or sets the marker background color as an index into the current color palette, or as one of the following XlColorIndex constants: xlColorIndexAutomatic or xlColorIndexNone. Read/write Long.
- `MarkerForegroundColor As Long  (read/write)`  
  Sets the marker foreground color as an RGB value or returns the corresponding color index value. Read/write Long.
- `MarkerForegroundColorIndex As XlColorIndex  (read/write)`  
  Returns or sets the marker foreground color as an index into the current color palette, or as one of the following XlColorIndex constants: xlColorIndexAutomatic or xlColorIndexNone. Read/write Long.
- `MarkerSize As Long  (read/write)`  
  Returns or sets the data-marker size, in points. Read/write Long.
- `MarkerStyle As XlMarkerStyle  (read/write)`  
  Returns or sets the marker style for a point or series in a line chart, scatter chart, or radar chart. Read/write XlMarkerStyle.
- `PictureType As Long  (read/write)`  
  Returns or sets the way pictures are displayed on a legend key. Read/write XlChartPictureType.
- `Smooth As Boolean  (read/write)`  
  True if curve smoothing is turned on for the legend key. Read/write Boolean.
- `Left As Double  (read-only)`  
  Returns the distance, in points, from the left edge of the object to the left edge of the chart area. Read-only Double.
- `Top As Double  (read-only)`  
  Returns the distance, in points, from the top edge of the object to the top of the first row (on a worksheet) or the top of the chart area (on a chart). Read-only Double.
- `Width As Double  (read-only)`  
  Returns the width, in points, of the object. Read-only Double.
- `Height As Double  (read-only)`  
  Returns the height, in points, of the object. Read-only Double.
- `Shadow As Boolean  (read/write)`  
  Returns or sets a value that indicates whether the object has a shadow. Read/write Boolean.
- `Format As ChartFormat  (read-only)`  
  Returns the line, fill, and effect formatting for the object. Read-only ChartFormat.
- `Application As Object  (read-only)`  
  When used without an object qualifier, returns an Application object that represents the Microsoft Word application. When used with an object qualifier, returns an Application object that represents the creator of the specified object (you can use this property with an Automation object to return the application of that object). Read-only.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `PictureUnit2 As Double  (read/write)`  
  Returns or sets the unit for each picture on the chart if the PictureType property is set to xlStackScale; otherwise, this property is ignored. Read/write Double.

## Methods (2)

- `ClearFormats() As Variant`  
  Clears the formatting of the object.
- `Delete() As Variant`  
  Deletes the object.
