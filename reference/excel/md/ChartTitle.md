# ChartTitle

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00020849-0000-0000-C000-000000000046}  

Represents the chart title.

**Remarks:** Use the ChartTitle property of the Chart object to return the ChartTitle object. The ChartTitle object doesn't exist and cannot be used unless the HasTitle property for the chart is True.

**Example:**

```vba
With Worksheets("sheet1").ChartObjects(1).Chart
 .HasTitle = True
 .ChartTitle.Text = "February Sales"
End With
```

## Properties (23)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Name As String  (read-only)`  
  Returns a String value that represents the name of the object.
- `Caption As String  (read/write)`  
  Returns or sets a String value that represents the chart title text.
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
- `IncludeInLayout As Boolean  (read/write)`  
  True if a chart title will occupy the chart layout space when a chart layout is being determined. The default value is True. Read/write Boolean.
- `Position As XlChartElementPosition  (read/write)`  
  Returns or sets the position of the chart title on the chart. Read/write XlChartElementPosition.
- `Format As ChartFormat  (read-only)`  
  Returns the ChartFormat object. Read-only.
- `Height As Double  (read-only)`  
  Returns the height, in points, of the object. Read-only.
- `Width As Double  (read-only)`  
  Returns the width, in points, of the object. Read-only.
- `Formula As String  (read/write)`  
  Gets or sets a String value that represents the formula of the object by using A1-style notation, in English. Read/write.
- `FormulaR1C1 As String  (read/write)`  
  Gets or sets a String value that represents the formula of the object by using R1C1-style notation, in English. Read/write.
- `FormulaLocal As String  (read/write)`  
  Gets or sets a String value that represents the formula of the object by using A1-style notation, in the language of the user. Read/write.
- `FormulaR1C1Local As String  (read/write)`  
  Gets or sets a String value that represents the formula of the object by using R1C1-style notation, in the language of the user. Read/write.

## Methods (4)

- `Select() As Variant`  
  Selects the object.
- `Delete() As Variant`  
  Deletes the object.
- `SetProperty(ID As String, Value As Variant)`
- `GetProperty(ID As String) As Variant`
