# ErrorBars

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {194F8476-B79D-4572-A609-294207DE77C1}  

Represents the error bars on a chart series.

**Remarks:** Error bars indicate the degree of uncertainty for chart data. Only series in area, bar, column, line, and scatter groups on a 2D chart can have error bars. Only series in scatter groups can have x and y error bars. This object is not a collection. There is no object that represents a single error bar; you either enable x error bars or y error bars for all points in a series or you disable them. The ErrorBar method changes the error bar format and type.

**Example:**

```vba
With ActiveDocument.InlineShapes(1)
 If .HasChart Then
 .Chart.SeriesCollection(1).HasErrorBars = True
 .Chart.SeriesCollection(1).ErrorBars.EndStyle = xlNoCap
 End If
End With
```

## Properties (7)

- `Parent As Object  (read-only)`  
  Returns the parent for the specified object. Read-only Object.
- `Name As String  (read-only)`  
  Returns the name of the object. Read-only String.
- `Border As ChartBorder  (read-only)`  
  Returns the border of the object. Read-only ChartBorder.
- `EndStyle As XlEndStyleCap  (read/write)`  
  Returns or sets the end style for the error bars. Read/write Long.
- `Format As ChartFormat  (read-only)`  
  Returns the line, fill, and effect formatting for the object. Read-only ChartFormat.
- `Application As Object  (read-only)`  
  When used without an object qualifier, returns an Application object that represents the Microsoft Word application. When used with an object qualifier, returns an Application object that represents the creator of the specified object (you can use this property with an Automation object to return the application of that object). Read-only.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.

## Methods (5)

- `Select() As Variant`  
  Selects the object.
- `Delete() As Variant`  
  Deletes the object.
- `ClearFormats() As Variant`  
  Clears the formatting of the object.
- `SetProperty(ID As String, Value As Variant)`
- `GetProperty(ID As String) As Variant`
