# DataLabel

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {1FD94DF1-3569-4465-94FF-E8B22D28EEB0}  

Represents the data label on a chart point or trendline.

**Remarks:** On a series, the DataLabel object is a member of the DataLabels collection. The DataLabels collection contains a DataLabel object for each point. For a series without definable points (such as an area series), the DataLabels collection contains a single DataLabel object.

**Example:**

```vba
With ActiveDocument.InlineShapes(1)
 If .HasChart Then
 .Chart.SeriesCollection(1).DataLabels(5).NumberFormat = "0.000"
 End If
End With
```

## Properties (34)

- `Parent As Object  (read-only)`  
  Returns the parent for the specified object. Read-only Object.
- `Name As String  (read-only)`  
  Returns the name of the object. Read-only String.
- `Caption As String  (read/write)`  
  Returns or sets the data label text. Read/write String.
- `Characters As ChartCharacters  (read-only)`  
  Returns a ChartCharacters object that represents a range of characters within the object text. Use the ChartCharacters object to format characters within a text string.
- `HorizontalAlignment As Variant  (read/write)`  
  Returns or sets the horizontal alignment for the specified object. Read/write Variant.
- `Left As Double  (read/write)`  
  Returns or sets the distance, in points, from the left edge of the object to the left edge of the chart area. Read/write Double.
- `Orientation As Variant  (read/write)`  
  Returns or sets the text orientation. Read/write Long.
- `Shadow As Boolean  (read/write)`  
  Returns or sets a value that indicates whether the object has a shadow. Read/write Boolean.
- `Text As String  (read/write)`  
  Returns or sets the text for the specified object. Read/write String.
- `Top As Double  (read/write)`  
  Returns or sets the distance, in points, from the top edge of the object to the top of the first row (on a worksheet) or the top of the chart area (on a chart). Read/write Double.
- `VerticalAlignment As Variant  (read/write)`  
  Returns or sets the vertical alignment of the specified object. Read/write Variant.
- `ReadingOrder As Long  (read/write)`  
  Returns or sets an XlReadingOrder constant that represents the reading order for the specified object. Read/write Long.
- `AutoText As Boolean  (read/write)`  
  True if the object automatically generates appropriate text based on context. Read/write Boolean.
- `NumberFormat As String  (read/write)`  
  Returns or sets the format code for the object. Read/write String.
- `NumberFormatLinked As Boolean  (read/write)`  
  True if the number format is linked to the cells (so that the number format changes in the labels when it changes in the cells). Read/write Boolean.
- `NumberFormatLocal As Variant  (read/write)`  
  Returns or sets the format code for the object as a string in the language of the user. Read/write Variant.
- `ShowLegendKey As Boolean  (read/write)`  
  True if the data label legend key is visible. Read/write Boolean.
- `Position As XlDataLabelPosition  (read/write)`  
  Returns or sets the position of the data label. Read/write XlDataLabelPosition.
- `ShowSeriesName As Boolean  (read/write)`  
  True to show the series name for the data labels on a chart. False to hide the series name. Read/write Boolean.
- `ShowCategoryName As Boolean  (read/write)`  
  True to display the category name for the data labels on a chart. False to hide the category name. Read/write Boolean.
- `ShowValue As Boolean  (read/write)`  
  True to display a specified chart's data label values. False to hide the values. Read/write Boolean.
- `ShowPercentage As Boolean  (read/write)`  
  True to display the percentage value for the data labels on a chart. False to hide the value. Read/write Boolean.
- `ShowBubbleSize As Boolean  (read/write)`  
  True to show the bubble size for the data labels on a chart. False to hide the bubble size. Read/write Boolean.
- `Separator As Variant  (read/write)`  
  Returns or sets the separator used for the data labels on a chart. Read/write Variant.
- `Format As ChartFormat  (read-only)`  
  Returns the line, fill, and effect formatting for the object. Read-only ChartFormat.
- `Application As Object  (read-only)`  
  When used without an object qualifier, returns an Application object that represents the Microsoft Word application. When used with an object qualifier, returns an Application object that represents the creator of the specified object (you can use this property with an Automation object to return the application of that object). Read-only.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Formula As String  (read/write)`  
  Gets or sets a String value that represents the formula of the object using A1-style notation, in English. Read/write.
- `FormulaR1C1 As String  (read/write)`  
  Returns or sets the formula for the object, using R1C1-style notation in the language of the macro. Read/write.
- `FormulaLocal As String  (read/write)`  
  Returns or sets the formula for the object, using A1-style references in the language of the user. Read/write String.
- `FormulaR1C1Local As String  (read/write)`  
  Returns or sets the formula for the object, using R1C1-style notation in the language of the user. Read/write String.
- `ShowRange As Boolean  (read/write)`  
  Set to True to display the Value From Cells range field for the specified chart data label. Set to False to hide that field. Read/write Boolean.
- `Height As Double  (read/write)`  
  Gets or sets the height, in points, of the object. Read/write Double.
- `Width As Double  (read/write)`  
  Gets or sets the width, in points, of the object. Read/write Double.

## Methods (4)

- `Select() As Variant`  
  Selects the object.
- `Delete() As Variant`  
  Deletes the object.
- `SetProperty(ID As String, Value As Variant)`
- `GetProperty(ID As String) As Variant`
