# LegendEntries

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {B6511068-70BF-4751-A741-55C1D41AD96F}  

A collection of all the LegendEntry objects in the specified chart legend.

**Remarks:** Each legend entry has two parts: - The text of the entry, which is the name of the series or trendline associated with the legend entry. - The entry marker, which visually links the legend entry with its associated series or trendline in the chart. The formatting properties for the entry marker and its associated series or trendline are contained in the LegendKey object.

**Example:**

```vba
With ActiveDocument.InlineShapes(1)
 If .HasChart Then
 With .Chart.Legend
 For i = 1 To .LegendEntries.Count
 .LegendEntries(i).Font.ColorIndex = 5
 Next
 End With
 End If
End With
```

## Properties (4)

- `Parent As Object  (read-only)`  
  Returns the parent for the specified object. Read-only Object.
- `Count As Long  (read-only)`  
  Returns the number of objects in the collection. Read-only Long.
- `Application As Object  (read-only)`  
  When used without an object qualifier, returns an Application object that represents the Microsoft Word application. When used with an object qualifier, returns an Application object that represents the creator of the specified object (you can use this property with an Automation object to return the application of that object). Read-only.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.

## Methods (3)

- `Item(Index As Variant) As LegendEntry`  
  Returns a single object from a collection.
    - `Index As Variant` (required): The index number for the object.
- `_NewEnum() As IUnknown`
- `_Default(Index As Variant) As LegendEntry`
