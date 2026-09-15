# ChartValuesCollection

**Type:** Dispatch Interface  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {11D6CD3F-7A25-47DF-A859-C1DC8CBF2620}  

A collection of all the ChartValues objects in the specified chart.

**Example:**

```vba
With myChart
 For Each cv In .ChartValuesCollection
  MsgBox (cv.Name)
 Next
End With
```

## Properties (4)

- `Application As Application  (read-only)`
- `Parent As Object  (read-only)`
- `Item As _ChartValues  (read-only)`
- `Count As Long  (read-only)`
