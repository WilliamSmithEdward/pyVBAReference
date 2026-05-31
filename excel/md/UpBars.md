# UpBars

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000208C5-0000-0000-C000-000000000046}  

Represents the up bars in a chart group.

**Remarks:** Up bars connect points on series one with higher values on the last series in the chart group (the lines go up from series one). Only 2D line groups that contain at least two series can have up bars. This object isn't a collection. There's no object that represents a single up bar; you either have up bars turned on for all points in a chart group or you have them turned off. If the HasUpDownBars property of the ChartGroup object is False, most properties of the UpBars object are disabled.

**Example:**

```vba
With Worksheets("sheet5").ChartObjects(1).Chart.ChartGroups(1)
 .HasUpDownBars = True
 .UpBars.Interior.Color = RGB(0, 0, 255)
 .DownBars.Interior.Color = RGB(255, 0, 0)
End With
```

## Properties (5)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Name As String  (read-only)`  
  Returns a String value that represents the name of the object.
- `Format As ChartFormat  (read-only)`  
  Returns the ChartFormat object. Read-only.

## Methods (4)

- `Select() As Variant`  
  Selects the object.
- `Delete() As Variant`  
  Deletes the object.
- `SetProperty(ID As String, Value As Variant)`
- `GetProperty(ID As String) As Variant`
