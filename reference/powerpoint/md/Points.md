# Points

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {92D41A74-F07E-4CA4-AF6F-BEF486AA4E6F}  

A collection of all the Point objects in the specified series in a chart.

**Remarks:** Use Points (_index_), where _index_ is the point index number, to return a single Point object. Points are numbered from left to right on the series. Points(1) is the leftmost point, and Points(Points.Count) is the rightmost point.

**Example:**

```vba
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.SeriesCollection(1).Points

            .Item(.Count).ApplyDataLabels Type:=xlShowValue

        End With

    End If

End With
```

## Properties (4)

- `Parent As Object  (read-only)`  
  Returns the parent for the specified object. Read-only Object.
- `Count As Long  (read-only)`  
  Returns the number of objects in the collection. Read-only Long.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Application As Application  (read-only)`  
  When used without an object qualifier, returns an Application object that represents the Microsoft PowerPoint application. When used with an object qualifier, returns an Application object that represents the creator of the specified object (you can use this property with an Automation object to return the application of that object). Read-only.

## Methods (3)

- `Item(Index As Long) As Point`  
  Returns a single object from a collection.
    - `Index As Long` (required): The index number for the object.
- `_NewEnum() As IUnknown`
- `_Default(Index As Long) As Point`
