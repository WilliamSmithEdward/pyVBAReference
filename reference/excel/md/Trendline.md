# Trendline

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000208BE-0000-0000-C000-000000000046}  

Represents a trendline in a chart.

**Remarks:** A trendline shows the trend, or direction, of data in a series. The Trendline object is a member of the Trendlines collection. The Trendlines collection contains all the Trendline objects for a single series.

**Example:**

```vba
Worksheets(1).ChartObjects(1).Chart. _
 SeriesCollection(1).Trendlines(1).Type = xlMovingAvg
```

## Properties (18)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Border As Border  (read-only)`  
  Returns a Border object that represents the border of the object.
- `DataLabel As DataLabel  (read-only)`  
  Returns a DataLabel object that represents the data label associated with the trendline. Read-only.
- `DisplayEquation As Boolean  (read/write)`  
  True if the equation for the trendline is displayed on the chart (in the same data label as the R-squared value). Setting this property to True automatically turns on data labels. Read/write Boolean.
- `DisplayRSquared As Boolean  (read/write)`  
  True if the R-squared value of the trendline is displayed on the chart (in the same data label as the equation). Setting this property to True automatically turns on data labels. Read/write Boolean.
- `Index As Long  (read-only)`  
  Returns a Long value that represents the index number of the object within the collection of similar objects.
- `Intercept As Double  (read/write)`  
  Returns or sets the point where the trendline crosses the value axis. Read/write Double.
- `InterceptIsAuto As Boolean  (read/write)`  
  True if the point where the trendline crosses the value axis is automatically determined by the regression. Read/write Boolean.
- `Name As String  (read/write)`  
  Returns or sets a String value representing the name of the object.
- `NameIsAuto As Boolean  (read/write)`  
  True if Microsoft Excel automatically determines the name of the trendline. Read/write Boolean.
- `Order As Long  (read/write)`  
  Returns or sets a Long value that represents the trendline order (an integer greater than 1) when the trendline type is xlPolynomial (XlTrendlineType).
- `Period As Long  (read/write)`  
  Returns or sets the period for the moving-average trendline. Can be a value from 2 through 255. Read/write Long.
- `Type As XlTrendlineType  (read/write)`  
  Returns or sets an XlTrendlineType value that represents the trendline type.
- `Backward2 As Double  (read/write)`  
  Returns or sets the number of periods (or units on a scatter chart) that the trendline extends backward. Read/write Double.
- `Forward2 As Double  (read/write)`  
  Returns or sets the number of periods (or units on a scatter chart) that the trendline extends forward. Read/write Double.
- `Format As ChartFormat  (read-only)`  
  Returns the ChartFormat object. Read-only.

## Methods (5)

- `ClearFormats() As Variant`  
  Clears the formatting of the object.
- `Delete() As Variant`  
  Deletes the object.
- `Select() As Variant`  
  Selects the object.
- `SetProperty(ID As String, Value As Variant)`
- `GetProperty(ID As String) As Variant`
