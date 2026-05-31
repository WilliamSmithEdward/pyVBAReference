# Legend

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000208CD-0000-0000-C000-000000000046}  

Represents the legend in a chart. Each chart can have only one legend.

**Remarks:** The Legend object contains one or more LegendEntry objects; each LegendEntry object contains a LegendKey object. The chart legend isn't visible unless the HasLegend property is True. If this property is False, properties and methods of the Legend object will fail.

**Example:**

```vba
Worksheets(1).ChartObjects(1).Chart.Legend.Font.Bold = True
```

## Properties (12)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Name As String  (read-only)`  
  Returns a String value that represents the name of the object.
- `Position As XlLegendPosition  (read/write)`  
  Returns or sets an XlLegendPosition value that represents the position of the legend on the chart.
- `Shadow As Boolean  (read/write)`  
  Returns or sets a Boolean value that determines if the object has a shadow.
- `Height As Double  (read/write)`  
  Returns or sets a Double value that represents the height, in points, of the object.
- `Left As Double  (read/write)`  
  Returns or sets a Double value that represents the distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).
- `Top As Double  (read/write)`  
  Returns or sets a Double value that represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).
- `Width As Double  (read/write)`  
  Returns or sets a Double value that represents the width, in points, of the object.
- `IncludeInLayout As Boolean  (read/write)`  
  True if a legend will occupy the chart layout space when a chart layout is being determined. The default value is True. Read/write Boolean.
- `Format As ChartFormat  (read-only)`  
  Returns the ChartFormat object. Read-only.

## Methods (6)

- `Select() As Variant`  
  Selects the object.
- `Delete() As Variant`  
  Deletes the object.
- `LegendEntries([Index As Variant]) As Object`  
  Returns an object that represents either a single legend entry (a LegendEntry object) or a collection of legend entries (a LegendEntries object) for the legend.
    - `Index As Variant` (optional): The number of the legend entry.
- `Clear() As Variant`  
  Clears the entire object.
- `SetProperty(ID As String, Value As Variant)`
- `GetProperty(ID As String) As Variant`
