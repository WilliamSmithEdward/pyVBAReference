# ChartData

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {92D41A5A-F07E-4CA4-AF6F-BEF486AA4E6F}  

Represents access to the linked or embedded data associated with a chart.

**Remarks:** Use the ChartData property to return the ChartData object.

**Example:**

```vba
With ActiveDocument.InlineShapes(1).Chart.ChartData

    .Activate

End With
```

## Properties (2)

- `Workbook As Object  (read-only)`  
  Returns the workbook that contains the chart data associated with the chart. Read-only Object.
- `IsLinked As Boolean  (read-only)`  
  True if the data for the chart is linked to an external Microsoft Excel workbook. Read-only Boolean.

## Methods (3)

- `Activate()`  
  Activates the first window of the workbook associated with the chart.
- `BreakLink()`  
  Removes the link between the data for a chart and a Microsoft Excel workbook.
- `ActivateChartDataWindow()`  
  Opens a Excel data grid window that contains the full source data for the specified chart.
