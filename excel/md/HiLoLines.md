# HiLoLines

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000208C2-0000-0000-C000-000000000046}  

Represents the high-low lines in a chart group.

**Remarks:** High-low lines connect the highest point with the lowest point in every category in the chart group. Only 2D line groups can have high-low lines. This object isn't a collection. There's no object that represents a single high-low line; you either have high-low lines turned on for all points in a chart group or you have them turned off. If the HasHiLoLines property of the ChartGroup object is False, most properties of the HiLoLines object are disabled.

**Example:**

```vba
Worksheets(1).ChartObjects(1).Activate
ActiveChart.ChartGroups(1).HasHiLoLines = True
ActiveChart.ChartGroups(1).HiLoLines.Border.Color = RGB(0, 0, 255)
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
