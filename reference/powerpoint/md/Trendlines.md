# Trendlines

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {92D41A7A-F07E-4CA4-AF6F-BEF486AA4E6F}  

Represents a collection of all the Trendline objects for the specified series.

**Remarks:** Each Trendline object represents a trendline in a chart. A trendline shows the trend, or direction, of data in a series.

**Example:**

```vba
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        MsgBox .Chart.SeriesCollection(1).Trendlines.Count

    End If

End With
```

## Properties (4)

- `Parent As Object  (read-only)`  
  Returns the parent for the specified object. Read-only Object.
- `Count As Long  (read-only)`  
  Returns the number of objects in the collection. Read-only Long.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Application As Application  (read-only)`  
  When used without an object qualifier, returns an Application object that represents the Microsoft PowerPoint application. When used with an object qualifier, returns an Application object that represents the creator of the specified object (you can use this property with an Automation object to return the application of that object). Read-only.

## Methods (4)

- `Add([Type As XlTrendlineType], [Order As Variant], [Period As Variant], [Forward As Variant], [Backward As Variant], [Intercept As Variant], [DisplayEquation As Variant], [DisplayRSquared As Variant], [Name As Variant]) As Trendline`  
  Creates a new trendline.
    - `Type As XlTrendlineType` (optional): One of the enumeration values that specifies the trendline type. The default is xlLinear.
    - `Order As Variant` (optional): The trendline order. Required ifType is set to xlPolynomial. If specified, the value must be an integer from 2 through 6.
    - `Period As Variant` (optional): The trendline period. Required ifType is set to xlMovingAvg. If specified, the value must be an integer greater than 1 and less than the number of data points in the series to which you are adding a trendline.
    - `Forward As Variant` (optional): The number of periods (or units on a scatter chart) that the trendline extends forward.
    - `Backward As Variant` (optional): The number of periods (or units on a scatter chart) that the trendline extends backward.
    - `Intercept As Variant` (optional): The trendline intercept. If specified, the value must be a double-precision floating-point number. If omitted, the intercept is automatically set by the regression, and the InterceptIsAuto property of the resulting Trendline object is set to True. This parameter is applicable only ifType is set to xlExponential, xlLinear, or xlPolynomial.
    - `DisplayEquation As Variant` (optional): True to display the equation of the trendline on the chart (in the same data label as the R-squared value). The default is False.
    - `DisplayRSquared As Variant` (optional): True to display the R-squared value of the trendline on the chart (in the same data label as the equation). The default is False.
    - `Name As Variant` (optional): The name of the trendline. If omitted, Microsoft Word generates a name, and the NameIsAuto property of the resulting Trendline object is set to True.
- `Item([Index As Variant]) As Trendline`  
  Returns a single object from a collection.
    - `Index As Variant` (optional): The index number for the object.
- `_NewEnum() As IUnknown`
- `_Default([Index As Variant]) As Trendline`
