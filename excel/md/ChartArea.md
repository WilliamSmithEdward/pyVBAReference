# ChartArea

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000208CC-0000-0000-C000-000000000046}  

Represents the chart area of a chart.

**Remarks:** The chart area includes everything, including the plot area. However, the plot area has its own fill, so filling the plot area does not fill the chart area. For information about formatting the plot area, see the PlotArea object. Use the ChartArea property of the Chart object to return the ChartArea object.

**Example:**

```vba
Worksheets("Sheet1").ChartObjects(1).Chart. _
 ChartArea.Format.Line.Visible = False
```

## Properties (11)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Name As String  (read-only)`  
  Returns a String value that represents the name of the object.
- `Shadow As Boolean  (read/write)`  
  Returns or sets a Boolean value that determines if the object has a shadow.
- `Height As Double  (read/write)`  
  Returns or sets a Double value that represents the height, in points, of the object.
- `Left As Double  (read/write)`  
  Returns or sets a Double value that represents the distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).
- `Top As Double  (read/write)`  
  Returns a Double value that represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).
- `Width As Double  (read/write)`  
  Returns or sets a Double value that represents the width, in points, of the object.
- `Format As ChartFormat  (read-only)`  
  Returns the ChartFormat object. Read-only.
- `RoundedCorners As Boolean  (read/write)`  
  True if the chart area of the chart has rounded corners. Read/write Boolean.

## Methods (5)

- `Select() As Variant`  
  Selects the object.
- `Clear() As Variant`  
  Clears the entire object.
- `Copy() As Variant`  
  Copies the object to the Clipboard.
- `ClearFormats() As Variant`  
  Clears the formatting of the object.
- `ClearContents() As Variant`  
  Clears the data from a chart but leaves the formatting.
