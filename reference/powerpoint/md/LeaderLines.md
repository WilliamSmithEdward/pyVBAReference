# LeaderLines

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {92D41A6D-F07E-4CA4-AF6F-BEF486AA4E6F}  

Represents leader lines on a chart. Leader lines connect data labels to data points.

**Remarks:** This object is not a collection; there is no object that represents a single leader line. This object applies only to pie charts.

**Example:**

```vba
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.SeriesCollection(1)

            .HasDataLabels = True

            .DataLabels.Position = xlLabelPositionBestFit

            .HasLeaderLines = True

            .LeaderLines.Border.ColorIndex = 5

        End With

    End If

End With
```

## Properties (5)

- `Border As ChartBorder  (read-only)`  
  Returns the border of the object. Read-only ChartBorder.
- `Format As ChartFormat  (read-only)`  
  Returns the line, fill, and effect formatting for the object. Read-only ChartFormat.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent for the specified object. Read-only Object.
- `Application As Application  (read-only)`  
  When used without an object qualifier, returns an Application object that represents the Microsoft PowerPoint application. When used with an object qualifier, returns an Application object that represents the creator of the specified object (you can use this property with an Automation object to return the application of that object). Read-only.

## Methods (2)

- `Select()`  
  Selects the object.
- `Delete()`  
  Deletes the object.
