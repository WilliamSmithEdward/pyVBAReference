# Legend

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {B3A1E8C6-E1CE-4A46-8D12-E017157B03D7}  

Represents the legend in a chart. Each chart can have only one legend.

**Remarks:** The Legend object contains one or more LegendEntry objects; each LegendEntry object contains a LegendKey object. The chart legend is not visible unless the HasLegend property is True. If this property is False, properties and methods of the Legend object will fail.

**Example:**

```vba
With ActiveDocument.InlineShapes(1)
 If .HasChart Then
 .Chart.Legend.Font.Bold = True
 End If
End With
```

## Properties (12)

- `Parent As Object  (read-only)`  
  Returns the parent for the specified object. Read-only Object.
- `Name As String  (read-only)`  
  Returns the name of the object. Read-only String.
- `Position As XlLegendPosition  (read/write)`  
  Returns or sets the position of the legend on the chart. Read/write XlLegendPosition.
- `Shadow As Boolean  (read/write)`  
  Returns or sets a value that indicates whether the object has a shadow. Read/write Boolean.
- `Height As Double  (read/write)`  
  Returns or sets the height, in points, of the object. Read/write Double.
- `Left As Double  (read/write)`  
  Returns the distance, in points, from the left edge of the object to the left edge of the chart area. Read-only Double.
- `Top As Double  (read/write)`  
  Returns or sets the distance, in points, from the top edge of the object to the top of the first row (on a worksheet) or the top of the chart area (on a chart). Read/write Double.
- `Width As Double  (read/write)`  
  Returns or sets the width, in points, of the object. Read/write Double.
- `IncludeInLayout As Boolean  (read/write)`  
  True if a legend will occupy the chart layout space when a chart layout is being determined. The default is True. Read/write Boolean.
- `Format As ChartFormat  (read-only)`  
  Returns the line, fill, and effect formatting for the object. Read-only ChartFormat.
- `Application As Object  (read-only)`  
  When used without an object qualifier, returns an Application object that represents the Microsoft Word application. When used with an object qualifier, returns an Application object that represents the creator of the specified object (you can use this property with an Automation object to return the application of that object). Read-only.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.

## Methods (6)

- `Select() As Variant`  
  Selects the object.
- `Delete() As Variant`  
  Deletes the object.
- `LegendEntries([Index As Variant]) As Object`  
  Returns a collection of legend entries for the legend.
- `Clear() As Variant`  
  Clears the entire object.
- `SetProperty(ID As String, Value As Variant)`
- `GetProperty(ID As String) As Variant`
