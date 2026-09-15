# DataLabels

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {D8252C5E-EB9F-4D74-AA72-C178B128FAC4}  

A collection of all the DataLabel objects for the specified series.

**Remarks:** Each DataLabel object represents a data label for a point or trendline. For a series without definable points (such as an area series), the DataLabels collection contains a single data label.

**Example:**

```vba
With ActiveDocument.InlineShapes(1)
 If .HasChart Then
 With Chart.SeriesCollection(1)
 .HasDataLabels = True
 .DataLabels.NumberFormat = "##.##"
 End With
 End If
End With
```

## Properties (24)

- `Parent As Object  (read-only)`  
  Returns the parent for the specified object. Read-only Object.
- `Name As String  (read-only)`  
  Returns the name of the object. Read-only String.
- `HorizontalAlignment As Variant  (read/write)`  
  Returns or sets the horizontal alignment for the specified object. Read/write Variant.
- `Orientation As Variant  (read/write)`  
  Returns or sets the text orientation. Read/write Long.
- `Shadow As Boolean  (read/write)`  
  Returns or sets a value that indicates whether the object has a shadow. Read/write Boolean.
- `VerticalAlignment As Variant  (read/write)`  
  Returns or sets the vertical alignment of the specified object. Read/write Variant.
- `ReadingOrder As Long  (read/write)`  
  Returns or sets an XlReadingOrder constant that represents the reading order for the specified object. Read/write Long.
- `AutoText As Boolean  (read/write)`  
  True if all objects in the collection automatically generate appropriate text based on context. Read/write Boolean.
- `NumberFormat As String  (read/write)`  
  Returns or sets the format code for the object. Read/write String.
- `NumberFormatLinked As Boolean  (read/write)`  
  True if the number format is linked to the cells (so that the number format changes in the labels when it changes in the cells). Read/write Boolean.
- `NumberFormatLocal As Variant  (read/write)`  
  Returns or sets the format code for the object as a string in the language of the user. Read/write Variant.
- `ShowLegendKey As Boolean  (read/write)`  
  True if the data label legend key is visible. Read/write Boolean.
- `Position As XlDataLabelPosition  (read/write)`  
  Returns or sets the position of the data labels. Read/write XlDataLabelPosition.
- `ShowSeriesName As Boolean  (read/write)`  
  True to show the series name for the data labels on a chart. False to hide the name. Read/write Boolean.
- `ShowCategoryName As Boolean  (read/write)`  
  True to display the category name for the data labels on a chart. False to hide the name. Read/write Boolean.
- `ShowValue As Boolean  (read/write)`  
  True to display the data label values for a specified chart. False to hide the values. Read/write Boolean.
- `ShowPercentage As Boolean  (read/write)`  
  True to display the percentage value for the data labels on a chart. False to hide the value. Read/write Boolean.
- `ShowBubbleSize As Boolean  (read/write)`  
  True to show the bubble size for the data labels on a chart. False to hide the bubble size. Read/write Boolean.
- `Separator As Variant  (read/write)`  
  Sets or returns the separator for the data labels on a chart. Read/write Variant.
- `Count As Long  (read-only)`  
  Returns the number of objects in the collection. Read-only Long.
- `Format As ChartFormat  (read-only)`  
  Returns the line, fill, and effect formatting for the object. Read-only ChartFormat.
- `Application As Object  (read-only)`  
  When used without an object qualifier, returns an Application object that represents the Microsoft Word application. When used with an object qualifier, returns an Application object that represents the creator of the specified object (you can use this property with an Automation object to return the application of that object). Read-only.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `ShowRange As Boolean  (read/write)`  
  Set to True to display the Value From Cells range field in all the chart data labels for a specified chart. Set to False to hide that field. Read/write Boolean.

## Methods (8)

- `Select() As Variant`  
  Selects the object.
- `Delete() As Variant`  
  Deletes the object.
- `Item(Index As Variant) As DataLabel`  
  Returns a single object from a collection.
    - `Index As Variant` (required): The index number for the object.
- `_NewEnum() As IUnknown`
- `_Default(Index As Variant) As DataLabel`
- `Propagate(Index As Variant)`  
  Propagates the contents and formatting of the specified data label to all the other data labels in the series.
    - `Index As Variant` (required): The index number in the DataLabels collection of the data label to propagate.
- `SetProperty(ID As String, Value As Variant)`
- `GetProperty(ID As String) As Variant`
