# Trendline

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {91C46192-3124-4346-A815-10B8873F5A06}  

Represents a trendline in a chart.

**Remarks:** A trendline shows the trend, or direction, of data in a series. The Trendline object is a member of the Trendlines collection. The Trendlines collection contains all the Trendline objects for a single series.

**Example:**

```vba
With ActiveDocument.InlineShapes(1)
 If .HasChart Then
 .Chart.SeriesCollection(1).Trendlines(1).Type = xlMovingAvg
 End If
End With
```

## Properties (18)

- `Parent As Object  (read-only)`  
  Returns the parent for the specified object. Read-only Object.
- `Border As ChartBorder  (read-only)`  
  Returns the border of the object. Read-only ChartBorder.
- `DataLabel As DataLabel  (read-only)`  
  Returns the data label that is associated with the trendline. Read-only DataLabel.
- `DisplayEquation As Boolean  (read/write)`  
  True if the equation for the trendline is displayed on the chart (in the same data label as the R-squared value). Read/write Boolean.
- `DisplayRSquared As Boolean  (read/write)`  
  True if the R-squared value of the trendline is displayed on the chart (in the same data label as the equation). Read/write Boolean.
- `Index As Long  (read-only)`  
  Returns the index number of the object within the collection of similar objects. Read-only Long.
- `Intercept As Double  (read/write)`  
  Returns or sets the point where the trendline crosses the value axis. Read/write Double.
- `InterceptIsAuto As Boolean  (read/write)`  
  True if the point where the trendline crosses the value axis is automatically determined by the regression. Read/write Boolean.
- `Name As String  (read/write)`  
  Returns or sets name of the object. Read/write String.
- `NameIsAuto As Boolean  (read/write)`  
  True if Microsoft Word automatically determines the name of the trendline. Read/write Boolean.
- `Order As Long  (read/write)`  
  Returns or sets the trendline order (an integer greater than 1) when the trendline type is xlPolynomial. Read/write Long.
- `Period As Long  (read/write)`  
  Returns or sets the period for the moving-average trendline. Read/write Long.
- `Type As XlTrendlineType  (read/write)`  
  Returns or sets the trendline type. Read/write XlTrendlineType.
- `Format As ChartFormat  (read-only)`  
  Returns the line, fill, and effect formatting for the object. Read-only ChartFormat.
- `Application As Object  (read-only)`  
  When used without an object qualifier, returns an Application object that represents the Microsoft Word application. When used with an object qualifier, returns an Application object that represents the creator of the specified object (you can use this property with an Automation object to return the application of that object). Read-only.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Backward2 As Double  (read/write)`  
  Returns or sets the number of periods (or units on a scatter chart) that the trendline extends backward. Read/write Double.
- `Forward2 As Double  (read/write)`  
  Returns or sets the number of periods (or units on a scatter chart) that the trendline extends forward. Read/write Double.

## Methods (5)

- `ClearFormats() As Variant`  
  Clears the formatting of the object.
- `Delete() As Variant`  
  Deletes the object.
- `Select() As Variant`  
  Selects the object.
- `SetProperty(ID As String, Value As Variant)`
- `GetProperty(ID As String) As Variant`
