# Gridlines

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000208C3-0000-0000-C000-000000000046}  

Represents major or minor gridlines on a chart axis.

**Remarks:** Gridlines extend the tick marks on a chart axis to make it easier to see the values associated with the data markers. This object isn't a collection. There's no object that represents a single gridline; you either have all gridlines for an axis turned on or all of them turned off. Use the MajorGridlines property of the Axis object to return the GridLines object that represents the major gridlines for the axis. Use the MinorGridlines property to return the GridLines object that represents the minor gridlines. It's possible to return both major and minor gridlines at the same time.

**Example:**

```vba
With Charts("chart1").Axes(xlCategory)
 .HasMajorGridlines = True
 .MajorGridlines.Border.Color = RGB(0, 0, 255)
 .MajorGridlines.Border.LineStyle = xlDash
End With
```

## Properties (6)

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
- `Format As ChartFormat  (read-only)`  
  Returns the ChartFormat object. Read-only.

## Methods (4)

- `Select() As Variant`  
  Selects the object.
- `Delete() As Variant`  
  Deletes the object.
- `SetProperty(ID As String, Value As Variant)`
- `GetProperty(ID As String) As Variant`
