# Gridlines

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {92D41A6A-F07E-4CA4-AF6F-BEF486AA4E6F}  

Represents major or minor gridlines on a chart axis.

**Remarks:** Gridlines extend the tick marks on a chart axis to make it easier to see the values associated with the data markers. This object is not a collection. There is no object that represents a single gridline; you either enable all gridlines for an axis or disable all of them. Use the MajorGridlines property to return the GridLines object that represents the major gridlines for the axis. Use the MinorGridlines property to return the GridLines object that represents the minor gridlines. It's possible to return both major and minor gridlines at the same time.

**Example:**

```vba
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.Axes(xlCategory)

            .HasMajorGridlines = True

            .MajorGridlines.Border.Color = RGB(0, 0, 255)

            .MajorGridlines.Border.LineStyle = xlDash

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
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Application As Application  (read-only)`  
  When used without an object qualifier, returns an Application object that represents the Microsoft PowerPoint application. When used with an object qualifier, returns an Application object that represents the creator of the specified object (you can use this property with an Automation object to return the application of that object). Read-only.

## Methods (4)

- `Select() As Variant`  
  Selects the object.
- `Delete() As Variant`  
  Deletes the object.
- `GetProperty(Id As String) As Variant`
- `SetProperty(Id As String, Value As Variant)`
