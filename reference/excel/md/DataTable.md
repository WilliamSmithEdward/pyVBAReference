# DataTable

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00020843-0000-0000-C000-000000000046}  

Represents a chart data table.

**Example:**

```vba
With Worksheets(1).ChartObjects(1).Chart
 .HasDataTable = True
 .DataTable.HasBorderOutline = True
End With
```

## Properties (10)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `ShowLegendKey As Boolean  (read/write)`  
  True if the data label legend key is visible. Read/write Boolean.
- `HasBorderHorizontal As Boolean  (read/write)`  
  True if the chart data table has horizontal cell borders. Read/write Boolean.
- `HasBorderVertical As Boolean  (read/write)`  
  True if the chart data table has vertical cell borders. Read/write Boolean.
- `HasBorderOutline As Boolean  (read/write)`  
  True if the chart data table has outline borders. Read/write Boolean.
- `Border As Border  (read-only)`  
  Returns a Border object that represents the border of the object.
- `Font As Font  (read-only)`  
  Returns a Font object that represents the font of the specified object.
- `Format As ChartFormat  (read-only)`  
  Returns the ChartFormat object. Read-only.

## Methods (2)

- `Select()`  
  Selects the object.
- `Delete()`  
  Deletes the object.
