# UpBars

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {86905AC9-33F3-4A88-96C8-B289B0390BCA}  

Represents the up bars in a chart group.

**Remarks:** Up bars connect points on series one with higher values on the last series in the chart group (the lines go up from series one). Only 2D line groups that contain at least two series can have up bars. This object is not a collection. There is no object that represents a single up bar; you either enable up bars for all points in a chart group or you disable them. If the HasUpDownBars property is False, most properties of the UpBars object are disabled.

**Example:**

```vba
With ActiveDocument.InlineShapes(1)
 If .HasChart Then
 With .Chart.ChartGroups(1)
 .HasUpDownBars = True
 .UpBars.Interior.Color = RGB(0, 0, 255)
 .DownBars.Interior.Color = RGB(255, 0, 0)
 End With
 End If
End With
```

## Properties (8)

- `Name As String  (read-only)`  
  Returns the name of the object. Read-only String.
- `Parent As Object  (read-only)`  
  Returns the parent for the specified object. Read-only Object.
- `Border As ChartBorder  (read-only)`  
  Returns the border of the object. Read-only ChartBorder.
- `Interior As Interior  (read-only)`  
  Returns the interior of the object. Read-only Interior.
- `Fill As ChartFillFormat  (read-only)`  
  Returns a FillFormat object for the parent chart element that contains fill formatting properties for the chart element. Read-only.
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
