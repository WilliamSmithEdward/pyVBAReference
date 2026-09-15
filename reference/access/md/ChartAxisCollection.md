# ChartAxisCollection

**Type:** Dispatch Interface  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {1657FD90-05DB-4B82-BA9C-ED9CDFC8A676}  

A collection of all the ChartAxis objects in the specified chart.

**Example:**

```vba
With myChart
 MsgBox (.ChartAxisCollection.Count)
  For Each axis In .ChartAxisCollection
    MsgBox (axis.Name)
  Next
End With
```

## Properties (4)

- `Application As Application  (read-only)`
- `Parent As Object  (read-only)`
- `Item As _ChartAxis  (read-only)`
- `Count As Long  (read-only)`
