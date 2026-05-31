# LeaderLines

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024437-0000-0000-C000-000000000046}  

Represents leader lines on a chart. Leader lines connect data labels to data points.

**Remarks:** This object isn't a collection; there's no object that represents a single leader line. This object applies only to pie charts.

**Example:**

```vba
With Worksheets(1).ChartObjects(1).Chart.SeriesCollection(1)
 .HasDataLabels = True
 .DataLabels.Position = xlLabelPositionBestFit
 .HasLeaderLines = True
 .LeaderLines.Border.ColorIndex = 5
End With
```

## Properties (5)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Border As Border  (read-only)`  
  Returns a Border object that represents the border of the object.
- `Format As ChartFormat  (read-only)`  
  Returns the ChartFormat object. Read-only.

## Methods (2)

- `Delete()`  
  Deletes the object.
- `Select()`  
  Selects the object.
