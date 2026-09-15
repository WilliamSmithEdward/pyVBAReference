# DropLines

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {92D41A66-F07E-4CA4-AF6F-BEF486AA4E6F}  

Represents the drop lines in a chart group.

**Remarks:** Drop lines connect the points in the chart with the x-axis. Only line and area chart groups can have drop lines. This object is not a collection. There is no object that represents a single drop line; you either enable drop lines for all points in a chart group or you disable them. If the HasDropLines property is False, most properties of the DropLines object are disabled.

**Example:**

```vba
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.ChartGroups(1)

            .HasDropLines = True

            .DropLines.Border.ColorIndex = 3

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

## Methods (2)

- `Delete() As Variant`  
  Deletes the object.
- `Select() As Variant`  
  Selects the object.
