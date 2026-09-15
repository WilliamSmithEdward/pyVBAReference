# DownBars

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {84A6A663-AEF4-4FCD-83FD-9BB707F157CA}  

Represents the down bars in a chart group.

**Remarks:** Down bars connect points on the first series in the chart group with lower values on the last series (the lines go down from the first series). Only 2D line groups that contain at least two series can have down bars. This object is not a collection. There is no object that represents a single down bar; you either enable up bars and down bars for all points in a chart group or you disable them. If the HasUpDownBars property is False, most properties of the DownBars object are disabled.

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

## Properties (5)

- `Name As String  (read-only)`  
  Returns the name of the object. Read-only String.
- `Parent As Object  (read-only)`  
  Returns the parent for the specified object. Read-only Object.
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
