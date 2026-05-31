# Trendlines

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000208BD-0000-0000-C000-000000000046}  

A collection of all the Trendline objects for the specified series.

**Remarks:** Each Trendline object represents a trendline in a chart. A trendline shows the trend, or direction, of data in a series.

**Example:**

```vba
MsgBox Charts(1).SeriesCollection(1).Trendlines.Count
```

## Properties (4)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Count As Long  (read-only)`  
  Returns a Long value that represents the number of objects in the collection.

## Methods (4)

- `Add([Type As XlTrendlineType], [Order As Variant], [Period As Variant], [Forward As Variant], [Backward As Variant], [Intercept As Variant], [DisplayEquation As Variant], [DisplayRSquared As Variant], [Name As Variant]) As Trendline`  
  Creates a new trendline.
    - `Type As XlTrendlineType` (optional): The trendline type.
    - `Order As Variant` (optional): Variant if _Type_ is xlPolynomial. The trendline order. Must be an integer from 2 to 6, inclusive.
    - `Period As Variant` (optional): Variant if _Type_ is xlMovingAvg. The trendline period. Must be an integer greater than 1 and less than the number of data points in the series you are adding a trendline to.
    - `Forward As Variant` (optional): The number of periods (or units on a scatter chart) that the trendline extends forward.
    - `Backward As Variant` (optional): The number of periods (or units on a scatter chart) that the trendline extends backward.
    - `Intercept As Variant` (optional): The trendline intercept. If this argument is omitted, the intercept is automatically set by the regression.
    - `DisplayEquation As Variant` (optional): True to display the equation of the trendline on the chart (in the same data label as the R-squared value). The default value is False.
    - `DisplayRSquared As Variant` (optional): True to display the R-squared value of the trendline on the chart (in the same data label as the equation). The default value is False.
    - `Name As Variant` (optional): The name of the trendline as text. If this argument is omitted, Microsoft Excel generates a name.
- `Item([Index As Variant]) As Trendline`  
  Returns a single object from a collection.
    - `Index As Variant` (optional): The index number for the object.
- `_NewEnum() As IUnknown`
- `_Default([Index As Variant]) As Trendline`
