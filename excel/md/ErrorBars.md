# ErrorBars

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000208CE-0000-0000-C000-000000000046}  

Represents the error bars in a chart series.

**Remarks:** Error bars indicate the degree of uncertainty for chart data. Only series in area, bar, column, line, and scatter groups on a 2D chart can have error bars. Only series in scatter groups can have x and y error bars. This object isn't a collection. There's no object that represents a single error bar; you either have x error bars or y error bars turned on for all points in a series or you have them turned off. The ErrorBar method of the Series object changes the error bar format and type.

**Example:**

```vba
Worksheets("sheet1").ChartObjects(1).Activate
ActiveChart.SeriesCollection(1).HasErrorBars = True
ActiveChart.SeriesCollection(1).ErrorBars.EndStyle = xlNoCap
```

## Properties (7)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Name As String  (read-only)`  
  Returns a String value that represents the name of the object.
- `Border As Border  (read-only)`  
  Returns a Border object that represents the border of the object.
- `EndStyle As XlEndStyleCap  (read/write)`  
  Returns or sets the end style for the error bars. Can be one of the following XlEndStyleCap constants: xlCap or xlNoCap. Read/write Long.
- `Format As ChartFormat  (read-only)`  
  Returns the ChartFormat object. Read-only.

## Methods (5)

- `Select() As Variant`  
  Selects the object.
- `Delete() As Variant`  
  Deletes the object.
- `ClearFormats() As Variant`  
  Clears the formatting of the object.
- `SetProperty(ID As String, Value As Variant)`
- `GetProperty(ID As String) As Variant`
