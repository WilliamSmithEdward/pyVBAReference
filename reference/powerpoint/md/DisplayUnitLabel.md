# DisplayUnitLabel

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {92D41A64-F07E-4CA4-AF6F-BEF486AA4E6F}  

Represents a unit label on an axis in the specified chart.

**Remarks:** Unit labels are useful for charting large values (for example, in the millions or billions). You can make the chart more readable by using a single unit label instead of large numbers at each tick mark.

**Example:**

```vba
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.Axes(xlValue)

            .DisplayUnit = xlMillions

            .HasDisplayUnitLabel = True

            With .DisplayUnitLabel

                .Caption = "Millions"

                .AutoScaleFont = False

            End With

        End With

    End If

End With
```

## Properties (22)

- `Caption As String  (read/write)`  
  Returns or sets the display unit label text. Read/write String.
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
- `Name As String  (read-only)`  
  Returns the name of the object. Read-only String.
- `Parent As Object  (read-only)`  
  Returns the parent for the specified object. Read-only Object.
- `Position As XlChartElementPosition  (read/write)`  
  Returns or sets the position of the unit label on an axis in the chart. Read/write XlChartElementPosition.
- `Format As ChartFormat  (read-only)`  
  Returns the line, fill, and effect formatting for the object. Read-only ChartFormat.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Application As Application  (read-only)`  
  When used without an object qualifier, returns an Application object that represents the Microsoft PowerPoint application. When used with an object qualifier, returns an Application object that represents the creator of the specified object (you can use this property with an Automation object to return the application of that object). Read-only.
- `ReadingOrder As Long  (read/write)`  
  Returns or sets an XlReadingOrder constant that represents the reading order for the specified object. Read/write Long.
- `Height As Double  (read-only)`  
  Returns the height, in points, of the object. Read-only.
- `Width As Double  (read-only)`  
  Returns the width, in points, of the object. Read-only.
- `Formula As String  (read/write)`  
  Returns or sets the object's formula in A1-style notation. Read/write.
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
- `GetProperty(Id As String) As Variant`
- `SetProperty(Id As String, Value As Variant)`
