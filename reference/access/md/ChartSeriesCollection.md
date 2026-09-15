# ChartSeriesCollection

**Type:** Dispatch Interface  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {312167E9-CE3B-4704-8221-59DE13631069}  

A collection of all the ChartSeries objects in the specified chart.

**Example:**

```vba
With myChart
 For Each series In .ChartSeriesCollection
  MsgBox (series.Name)
 Next
End With
```

## Properties (4)

- `Application As Application  (read-only)`
- `Parent As Object  (read-only)`
- `Item As _ChartSeries  (read-only)`
- `Count As Long  (read-only)`
