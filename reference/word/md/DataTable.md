# DataTable

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {DCE9F2C4-4C02-43BA-840E-B4276550EF79}  

Represents a chart data table.

**Example:**

```vba
With ActiveDocument.InlineShapes(1)
 If .HasChart Then
 .Chart.HasDataTable = True
 .Chart.DataTable.HasBorderOutline = True
 End If
End With
```

## Properties (10)

- `ShowLegendKey As Boolean  (read/write)`  
  True if the data label legend key is visible. Read/write Boolean.
- `HasBorderHorizontal As Boolean  (read/write)`  
  True if the chart data table has horizontal cell borders. Read/write Boolean.
- `HasBorderVertical As Boolean  (read/write)`  
  True if the chart data table has vertical cell borders. Read/write Boolean.
- `HasBorderOutline As Boolean  (read/write)`  
  True if the chart data table has outline borders. Read/write Boolean.
- `Border As ChartBorder  (read-only)`  
  Returns the border of the object. Read-only ChartBorder.
- `Font As ChartFont  (read-only)`  
  Returns the font of the specified object. Read-only ChartFont.
- `Parent As Object  (read-only)`  
  Returns the parent for the specified object. Read-only Object.
- `Format As ChartFormat  (read-only)`  
  Returns the line, fill, and effect formatting for the object. Read-only ChartFormat.
- `Application As Object  (read-only)`  
  When used without an object qualifier, returns an Application object that represents the Microsoft Word application. When used with an object qualifier, returns an Application object that represents the creator of the specified object (you can use this property with an Automation object to return the application of that object). Read-only.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.

## Methods (2)

- `Select()`  
  Selects the object.
- `Delete()`  
  Deletes the object.
