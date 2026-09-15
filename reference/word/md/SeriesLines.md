# SeriesLines

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {5D7F6C15-36CE-44CC-9692-5A1F8B8C906D}  

Represents series lines in a chart group.

**Remarks:** Series lines connect the data values from each series. Only 2D stacked bar, 2D stacked column, pie-of-pie, or bar-of-pie charts can have series lines. This object is not a collection. There is no object that represents a single series line; you either enable series lines for all points in a chart group or you disable them. If the HasSeriesLines property is False, most properties of the SeriesLines object are disabled.

**Example:**

```vba
With ActiveDocument.InlineShapes(1)
 If .HasChart Then
 With .Chart.ChartGroups(1)
 .HasSeriesLines = True
 .SeriesLines.Border.Color = RGB(0, 0, 255)
 End With
 End If
End With
```

## Properties (6)

- `Name As String  (read-only)`  
  Returns the name of the object. Read-only String.
- `Parent As Object  (read-only)`  
  Returns the parent for the specified object. Read-only Object.
- `Border As ChartBorder  (read-only)`  
  Returns the border of the object. Read-only ChartBorder.
- `Format As ChartFormat  (read-only)`  
  Returns the line, fill, and effect formatting for the object. Read-only ChartFormat.
- `Application As Object  (read-only)`  
  When used without an object qualifier, returns an Application object that represents the Microsoft Word application. When used with an object qualifier, returns an Application object that represents the creator of the specified object (you can use this property with an Automation object to return the application of that object). Read-only.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.

## Methods (4)

- `Select() As Variant`  
  Selects the object.
- `Delete() As Variant`  
  Deletes the object.
- `SetProperty(ID As String, Value As Variant)`
- `GetProperty(ID As String) As Variant`
