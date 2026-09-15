# ChartArea

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {C75AD98A-74E9-49FE-8BF1-544839CC08A5}  

Represents the chart area of a chart.

**Remarks:** The chart area includes everything, including the plot area. However, the PlotArea object has its own formatting, so formatting the plot area does not format the chart area. Use the ChartArea property to return the ChartArea object.

**Example:**

```vba
With ActiveDocument.InlineShapes(1).Chart
 ChartArea.Format.Line.Visible = False
End With
```

## Properties (10)

- `Name As String  (read-only)`  
  Returns the name of the object. Read-only String.
- `Parent As Object  (read-only)`  
  Returns the parent for the specified object. Read-only Object.
- `Shadow As Boolean  (read/write)`  
  Returns or sets a value that determines whether the object has a shadow. Read/write Boolean.
- `Height As Double  (read/write)`  
  Returns or sets the height, in points, of the object. Read/write Double.
- `Left As Double  (read/write)`  
  Returns or sets the distance, in points, from the left edge of the object to the left edge of the chart area. Read/write Double.
- `Top As Double  (read/write)`  
  Returns or sets the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart). Read/write Double.
- `Width As Double  (read/write)`  
  Returns or sets the width, in points, of the object. Read/write Double.
- `Format As ChartFormat  (read-only)`  
  Returns the line, fill, and effect formatting for the object. Read-only ChartFormat.
- `Application As Object  (read-only)`  
  When used without an object qualifier, returns an Application object that represents the Microsoft Word application. When used with an object qualifier, returns an Application object that represents the creator of the specified object (you can use this property with an Automation object to return the application of that object). Read-only.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.

## Methods (5)

- `Select() As Variant`  
  Selects the object.
- `Clear() As Variant`  
  Clears the entire object.
- `ClearContents() As Variant`  
  Clears the data from a chart but leaves the formatting.
- `Copy() As Variant`  
  Copies the object to the Clipboard.
- `ClearFormats() As Variant`  
  Clears the formatting of the object.
