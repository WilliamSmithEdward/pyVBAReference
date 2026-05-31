# Points

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00020869-0000-0000-C000-000000000046}  

A collection of all the Point objects in the specified series in a chart.

**Remarks:** Use Points (_index_), where _index_ is the point index number, to return a single Point object. Points are numbered from left to right on the series. Points(1) is the leftmost point, and Points(Points.Count) is the rightmost point.

**Example:**

```vba
Dim pts As Points
Set pts = Worksheets(1).ChartObjects(1).Chart. _
 SeriesCollection(1).Points
pts(pts.Count).ApplyDataLabels type:=xlShowValue
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

- `Item(Index As Long) As Point`  
  Returns a single object from a collection.
    - `Index As Long` (required): The index number for the object.
- `_NewEnum() As IUnknown`
- `_Default(Index As Long) As Point`
