# DataLabel

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000208B2-0000-0000-C000-000000000046}  

Represents the data label on a chart point or trendline.

**Remarks:** In a series, the DataLabel object is a member of the DataLabels collection. The DataLabels collection contains a DataLabel object for each point. For a series without definable points (such as an area series), the DataLabels collection contains a single DataLabel object.

**Example:**

```vba
Worksheets(1).ChartObjects(1).Chart _
 .SeriesCollection(1).DataLabels(5).NumberFormat = "0.000"
```

## Properties (34)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Name As String  (read-only)`  
  Returns a String value that represents the name of the object.
- `Caption As String  (read/write)`  
  Returns or sets a String value that represents the data label text.
- `Characters As Characters  (read-only)`  
  Returns a Characters object that represents a range of characters within the object text. Use the Characters object to format characters within a text string.
- `HorizontalAlignment As Variant  (read/write)`  
  Returns or sets a Variant value that represents the horizontal alignment for the specified object.
- `Left As Double  (read/write)`  
  Returns or sets a Double value that represents the distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).
- `Orientation As Variant  (read/write)`  
  Returns or sets a Variant value that represents the text orientation.
- `Shadow As Boolean  (read/write)`  
  Returns or sets a Boolean value that determines if the object has a shadow.
- `Text As String  (read/write)`  
  Returns or sets the text for the specified object. Read/write String.
- `Top As Double  (read/write)`  
  Returns or sets a Double value that represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).
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
- `Format As ChartFormat  (read-only)`  
  Returns the ChartFormat object. Read-only.
- `Formula As String  (read/write)`  
  Gets or sets a String value that represents the formula of the object using A1-style notation, in English. Read/write.
- `FormulaR1C1 As String  (read/write)`  
  Gets or sets a String value that represents the formula of the object using R1C1-style notation, in English. Read/write.
- `FormulaLocal As String  (read/write)`  
  Gets or sets a String value that represents the formula of the object using A1-style notation, in the language of the user. Read/write.
- `FormulaR1C1Local As String  (read/write)`  
  Gets or sets a String value that represents the formula of the object using R1C1-style notation, in the language of the user. Read/write.
- `ShowRange As Boolean  (read/write)`  
  Used to toggle the Range field if it exists on the data label range. Read/write Boolean.
- `Height As Double  (read/write)`  
  Returns the height of the object in points. Read/write.
- `Width As Double  (read/write)`  
  Returns the width of the object in points. Read-only.

## Methods (4)

- `Select() As Variant`  
  Selects the object.
- `Delete() As Variant`  
  Deletes the object.
- `SetProperty(ID As String, Value As Variant)`
- `GetProperty(ID As String) As Variant`
