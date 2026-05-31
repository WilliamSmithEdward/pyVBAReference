# DropLines

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000208C4-0000-0000-C000-000000000046}  

Represents the drop lines in a chart group.

**Remarks:** Drop lines connect the points in the chart with the x-axis. Only line and area chart groups can have drop lines. This object isn't a collection. There's no object that represents a single drop line; you either have drop lines turned on for all points in a chart group or you have them turned off. If the HasDropLines property of the ChartGroup object is False, most properties of the DropLines object are disabled.

**Example:**

```vba
Worksheets("sheet1").ChartObjects(1).Activate
ActiveChart.ChartGroups(1).HasDropLines = True
ActiveChart.ChartGroups(1).DropLines.Border.ColorIndex = 3
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

## Methods (2)

- `Select() As Variant`  
  Selects the object.
- `Delete() As Variant`  
  Deletes the object.
