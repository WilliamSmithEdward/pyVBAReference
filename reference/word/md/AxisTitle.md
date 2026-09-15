# AxisTitle

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {8B0E45DB-3A7B-42EE-9D17-A92AF69B79C1}  

Represents a chart axis title.

**Remarks:** Use the AxisTitle property to return an AxisTitle object. The AxisTitle object does not exist and cannot be used unless the HasTitle property for the axis is True.

**Example:**

```vba
With ActiveDocument.InlineShapes(1)
 If .HasChart Then
 With .Chart.Axes(xlValue)
 .HasTitle = True
 With .AxisTitle
 .Caption = "Revenue (millions)"
 .Font.Name = "bookman"
 .Font.Size = 10
 .Characters(10, 8).Font.Italic = True
 End With
 End With
 End If
End With
```

## Properties (23)

- `Caption As String  (read/write)`  
  Returns or sets a String value that represents the axis title text.
- `Characters As ChartCharacters  (read-only)`  
  Returns a ChartCharacters object that represents a range of characters within the object text. Use the ChartCharacters object to format characters within a text string.
- `HorizontalAlignment As Variant  (read/write)`  
  Returns or sets the horizontal alignment for the specified object. Read/write Variant.
- `Left As Double  (read/write)`  
  Returns or sets the distance, in points, from the left edge of the object to the left edge of the chart area. Read/write Double.
- `Orientation As Variant  (read/write)`  
  Returns or sets the text orientation. Read/write Long.
- `Shadow As Boolean  (read/write)`  
  Returns or sets a value that determines whether the object has a shadow. Read/write Boolean.
- `Text As String  (read/write)`  
  Returns or sets the text for the specified object. Read/write String.
- `Top As Double  (read/write)`  
  Returns or sets the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart). Read/write Double.
- `VerticalAlignment As Variant  (read/write)`  
  Returns or sets the vertical alignment of the specified object. Read/write Variant.
- `ReadingOrder As Long  (read/write)`  
  Returns or sets an XlReadingOrder constant that represents the reading order for the specified object. Read/write Long.
- `Name As String  (read-only)`  
  Returns a String value that represents the name of the object.
- `Parent As Object  (read-only)`  
  Returns the parent for the specified object. Read-only Object.
- `IncludeInLayout As Boolean  (read/write)`  
  True if an axis title will occupy the chart layout space when a chart layout is being determined. The default is True. Read/write Boolean.
- `Position As XlChartElementPosition  (read/write)`  
  Returns or sets the position of the axis title on the chart. Read/write XlChartElementPosition.
- `Format As ChartFormat  (read-only)`  
  Returns the line, fill, and effect formatting for the object. Read-only ChartFormat.
- `Application As Object  (read-only)`  
  When used without an object qualifier, returns an Application object that represents the Microsoft Word application. When used with an object qualifier, returns an Application object that represents the creator of the specified object (you can use this property with an Automation object to return the application of that object). Read-only.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Height As Double  (read-only)`  
  Returns the height, in points, of the object. Read-only.
- `Width As Double  (read-only)`  
  Returns the width, in points, of the object. Read-only.
- `Formula As String  (read/write)`  
  Gets or sets a String value that represents the formula of the object using A1-style notation, in English. Read/write.
- `FormulaR1C1 As String  (read/write)`  
  Returns or sets the formula for the object, using R1C1-style notation in the language of the macro. Read/write String.
- `FormulaLocal As String  (read/write)`  
  Returns or sets the formula for the object, using A1-style references in the language of the user. Read/write String.
- `FormulaR1C1Local As String  (read/write)`  
  Returns or sets the formula for the object, using R1C1-style notation in the language of the user. Read/write String.

## Methods (4)

- `Delete() As Variant`  
  Deletes the object.
- `Select() As Variant`  
  Selects the object.
- `SetProperty(ID As String, Value As Variant)`
- `GetProperty(ID As String) As Variant`
