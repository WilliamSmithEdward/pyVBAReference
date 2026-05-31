# LegendEntries

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000208BB-0000-0000-C000-000000000046}  

A collection of all the LegendEntry objects in the specified chart legend.

**Remarks:** Each legend entry has two parts: the text of the entry, which is the name of the series or trendline associated with the legend entry; and the entry marker, which visually links the legend entry with its associated series or trendline in the chart. The formatting properties for the entry marker and its associated series or trendline are contained in the LegendKey object.

**Example:**

```vba
With Worksheets("sheet1").ChartObjects(1).Chart.Legend
 For i = 1 To .LegendEntries.Count
 .LegendEntries(i).Font.ColorIndex = 5
 Next
End With
```

## Properties (4)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Count As Long  (read-only)`  
  Returns a Long value that represents the number of objects in the collection.

## Methods (3)

- `Item(Index As Variant) As LegendEntry`  
  Returns a single object from a collection.
    - `Index As Variant` (required): The index number for the object.
- `_NewEnum() As IUnknown`
- `_Default(Index As Variant) As LegendEntry`
