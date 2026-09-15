# TickLabels

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {935D59F5-E365-4F92-B7F5-1C499A63ECA8}  

Represents the tick-mark labels associated with tick marks on a chart axis.

**Remarks:** This object is not a collection. There is no object that represents a single tick-mark label; you must return all the tick-mark labels as a unit. Tick-mark label text for the category axis comes from the name of the associated category in the chart. The default tick-mark label text for the category axis is the number that indicates the position of the category relative to the left end of this axis. To change the number of unlabeled tick marks between tick-mark labels, you must change the TickLabelSpacing property for the category axis. Tick-mark label text for the value axis is calculated based on the MajorUnit, MinimumScale, and MaximumScale properties of the value axis. To change the tick-mark label text for the value axis, you must change the values of these properties.

**Example:**

```vba
With ActiveDocument.InlineShapes(1)
 If .HasChart Then
 .Chart.Axes(xlValue).TickLabels.NumberFormat = "0.00"
 End If
End With
```

## Properties (15)

- `Parent As Object  (read-only)`  
  Returns the parent for the specified object. Read-only Object.
- `Font As ChartFont  (read-only)`  
  Returns the font of the specified object. Read-only ChartFont.
- `Name As String  (read-only)`  
  Returns the name of the object. Read-only String.
- `NumberFormat As String  (read/write)`  
  Returns or sets the format code for the object. Read/write String.
- `NumberFormatLinked As Boolean  (read/write)`  
  True if the number format is linked to the cells (so that the number format changes in the labels when it changes in the cells). Read/write Boolean.
- `NumberFormatLocal As Variant  (read/write)`  
  Returns or sets the format code for the object as a string in the language of the user. Read/write Variant.
- `Orientation As XlTickLabelOrientation  (read/write)`  
  Returns or sets the text orientation. Read/write Long.
- `ReadingOrder As Long  (read/write)`  
  Returns or sets an XlReadingOrder constant that represents the reading order for the specified object. Read/write Long.
- `Depth As Long  (read-only)`  
  Returns the number of levels of category tick labels. Read-only Long.
- `Offset As Long  (read/write)`  
  Returns or sets the distance between the levels of labels, and the distance between the first level and the axis line. Read/write Long.
- `Alignment As Long  (read/write)`  
  Returns or sets the alignment for the specified phonetic text or tick label. Read/write Long.
- `MultiLevel As Boolean  (read/write)`  
  Returns or sets a value that indicates whether an axis is multilevel. Read/write Boolean.
- `Format As ChartFormat  (read-only)`  
  Returns the line, fill, and effect formatting for the object. Read-only ChartFormat.
- `Application As Object  (read-only)`  
  When used without an object qualifier, returns an Application object that represents the Microsoft Word application. When used with an object qualifier, returns an Application object that represents the creator of the specified object (you can use this property with an Automation object to return the application of that object). Read-only.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.

## Methods (2)

- `Delete() As Variant`  
  Deletes the object.
- `Select() As Variant`  
  Selects the object.
