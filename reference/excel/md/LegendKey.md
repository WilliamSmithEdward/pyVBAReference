# LegendKey

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000208BC-0000-0000-C000-000000000046}  

Represents a legend key in a chart legend.

**Remarks:** Each legend key is a graphic that visually links a legend entry with its associated series or trendline in the chart. The legend key is linked to its associated series or trendline in such a way that changing the formatting of one simultaneously changes the formatting of the other.

**Example:**

```vba
Worksheets("sheet1").ChartObjects(1).Chart _
 .Legend.LegendEntries(1).LegendKey.MarkerBackgroundColorIndex = 5
```

## Properties (19)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `InvertIfNegative As Boolean  (read/write)`  
  True if Microsoft Excel inverts the pattern in the item when it corresponds to a negative number. Read/write Boolean.
- `MarkerBackgroundColor As Long  (read/write)`  
  Sets the marker background color as an RGB value or returns the corresponding color index value. Applies only to line, scatter, and radar charts. Read/write Long.
- `MarkerBackgroundColorIndex As XlColorIndex  (read/write)`  
  Returns or sets the marker background color as an index into the current color palette, or as one of the following XlColorIndex constants: xlColorIndexAutomatic or xlColorIndexNone. Applies only to line, scatter, and radar charts. Read/write Long.
- `MarkerForegroundColor As Long  (read/write)`  
  Sets the marker foreground color as an RGB value or returns the corresponding color index value. Applies only to line, scatter, and radar charts. Read/write Long.
- `MarkerForegroundColorIndex As XlColorIndex  (read/write)`  
  Returns or sets the marker foreground color as an index into the current color palette, or as one of the following XlColorIndex constants: xlColorIndexAutomatic or xlColorIndexNone. Applies only to line, scatter, and radar charts. Read/write Long.
- `MarkerSize As Long  (read/write)`  
  Returns or sets the data-marker size, in points. Can be a value from 2 through 72. Read/write Long.
- `MarkerStyle As XlMarkerStyle  (read/write)`  
  Returns or sets the marker style for a point or series in a line chart, scatter chart, or radar chart. Read/write XlMarkerStyle.
- `PictureType As Long  (read/write)`  
  Returns or sets an XlChartPictureType value that represents the way pictures are displayed on a legend key.
- `Smooth As Boolean  (read/write)`  
  True if curve smoothing is turned on for the legend key. Read/write.
- `Left As Double  (read-only)`  
  Returns a Double value that represents the distance, in points, from the left edge of the object to the left edge of the chart area.
- `Top As Double  (read-only)`  
  Returns a Double value that represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).
- `Width As Double  (read-only)`  
  Returns a Double value that represents the width, in points, of the object.
- `Height As Double  (read-only)`  
  Returns a Double value that represents the height, in points, of the object.
- `Shadow As Boolean  (read/write)`  
  Returns or sets a Boolean value that determines if the object has a shadow.
- `PictureUnit2 As Double  (read/write)`  
  Returns or sets the unit for each picture on the chart if the PictureType property is set to xlStackScale (if not, this property is ignored). Read/write Double.
- `Format As ChartFormat  (read-only)`  
  Returns the ChartFormat object. Read-only.

## Methods (2)

- `ClearFormats() As Variant`  
  Clears the formatting of the object.
- `Delete() As Variant`  
  Deletes the object.
