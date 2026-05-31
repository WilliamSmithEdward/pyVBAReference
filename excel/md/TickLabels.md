# TickLabels

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000208C9-0000-0000-C000-000000000046}  

Represents the tick-mark labels associated with tick marks on a chart axis.

**Remarks:** This object isn't a collection. There's no object that represents a single tick-mark label; you must return all the tick-mark labels as a unit. Tick-mark label text for the category axis comes from the name of the associated category in the chart. The default tick-mark label text for the category axis is the number that indicates the position of the category relative to the left end of this axis. To change the number of unlabeled tick marks between tick-mark labels, you must change the TickLabelSpacing property for the category axis. Tick-mark label text for the value axis is calculated based on the MajorUnit, MinimumScale, and MaximumScale properties of the value axis. To change the tick-mark label text for the value axis, you must change the values of these properties.

**Example:**

```vba
Worksheets("sheet1").ChartObjects(1).Chart _
 .Axes(xlValue).TickLabels.NumberFormat = "0.00"
```

## Properties (15)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Font As Font  (read-only)`  
  Returns a Font object that represents the font of the specified object.
- `Name As String  (read-only)`  
  Returns a String value that represents the name of the object.
- `NumberFormat As String  (read/write)`  
  Returns or sets a String value that represents the format code for the object.
- `NumberFormatLinked As Boolean  (read/write)`  
  True if the number format is linked to the cells (so that the number format changes in the labels when it changes in the cells). Read/write Boolean.
- `NumberFormatLocal As Variant  (read/write)`  
  Returns or sets a Variant value that represents the format code for the object as a string in the language of the user.
- `Orientation As XlTickLabelOrientation  (read/write)`  
  Returns or sets a Long value that represents the text orientation.
- `ReadingOrder As Long  (read/write)`  
  Returns or sets the reading order for the specified object. Can be one of the following XlReadingOrder constants: xlRTL (right-to-left), xlLTR (left-to-right), or xlContext. Read/write Long.
- `Depth As Long  (read-only)`  
  Returns a Long value that represents the number of levels of category tick labels.
- `Offset As Long  (read/write)`  
  Returns or sets a Long value that represents the distance between the levels of labels, and the distance between the first level and the axis line.
- `Alignment As Long  (read/write)`  
  Returns or sets a Long value that represents the alignment for the specified phonetic text or tick label.
- `MultiLevel As Boolean  (read/write)`  
  Sets whether an axis is multilevel. Read/write Boolean.
- `Format As ChartFormat  (read-only)`  
  Returns the ChartFormat object. Read-only.

## Methods (2)

- `Delete() As Variant`  
  Deletes the object.
- `Select() As Variant`  
  Selects the object.
