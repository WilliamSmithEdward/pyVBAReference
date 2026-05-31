# DataLabels

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000208B3-0000-0000-C000-000000000046}  

A collection of all the DataLabel objects for the specified series.

**Remarks:** Each DataLabel object represents a data label for a point or trendline. For a series without definable points (such as an area series), the DataLabels collection contains a single data label.

**Example:**

```vba
With Charts(1).SeriesCollection(1)
 .HasDataLabels = True
 .DataLabels.NumberFormat = "##.##"
End With
```

## Properties (24)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Name As String  (read-only)`  
  Returns a String value that represents the name of the object.
- `HorizontalAlignment As Variant  (read/write)`  
  Returns or sets a Variant value that represents the horizontal alignment for the specified object.
- `Orientation As Variant  (read/write)`  
  Returns or sets a Variant value that represents the text orientation.
- `Shadow As Boolean  (read/write)`  
  Returns or sets a Boolean value that determines if the object has a shadow.
- `VerticalAlignment As Variant  (read/write)`  
  Returns or sets a Variant value that represents the vertical alignment of the specified object.
- `ReadingOrder As Long  (read/write)`  
  Returns or sets the reading order for the specified object. Can be one of the following XlReadingOrder constants: xlRTL (right-to-left), xlLTR (left-to-right), or xlContext. Read/write Long.
- `AutoText As Boolean  (read/write)`  
  True if the object automatically generates appropriate text based on context. Read/write Boolean.
- `NumberFormat As String  (read/write)`  
  Returns or sets a String value that represents the format code for the object.
- `NumberFormatLinked As Boolean  (read/write)`  
  True if the number format is linked to the cells (so that the number format changes in the labels when it changes in the cells). Read/write Boolean.
- `NumberFormatLocal As Variant  (read/write)`  
  Returns or sets a Variant value that represents the format code for the object as a string in the language of the user.
- `ShowLegendKey As Boolean  (read/write)`  
  True if the data label legend key is visible. Read/write Boolean.
- `Position As XlDataLabelPosition  (read/write)`  
  Returns or sets an XlDataLabelPosition value that represents the position of the data label.
- `ShowSeriesName As Boolean  (read/write)`  
  Returns or sets a Boolean to indicate the series name display behavior for the data labels on a chart. True to show the series name. False to hide. Read/write.
- `ShowCategoryName As Boolean  (read/write)`  
  True to display the category name for the data labels on a chart. False to hide. Read/write Boolean.
- `ShowValue As Boolean  (read/write)`  
  Returns or sets a Boolean corresponding to a specified chart's data label values display behavior. True displays the values. False to hide. Read/write.
- `ShowPercentage As Boolean  (read/write)`  
  True to display the percentage value for the data labels on a chart. False to hide. Read/write Boolean.
- `ShowBubbleSize As Boolean  (read/write)`  
  True to show the bubble size for the data labels on a chart. False to hide. Read/write Boolean.
- `Separator As Variant  (read/write)`  
  Sets or returns a Variant representing the separator used for the data labels on a chart. Read/write.
- `Count As Long  (read-only)`  
  Returns a Long value that represents the number of objects in the collection.
- `Format As ChartFormat  (read-only)`  
  Returns the ChartFormat object. Read-only.
- `ShowRange As Boolean  (read/write)`  
  Used to toggle the Range field if it exists on the data label range. Read/write Boolean.

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
  Enables you to take the contents and formatting of a single data label and apply it to every other data label in the series.
    - `Index As Variant` (required): The index number of the data label to propagate.
- `SetProperty(ID As String, Value As Variant)`
- `GetProperty(ID As String) As Variant`
