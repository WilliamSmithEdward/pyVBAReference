# ColorStops

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244AE-0000-0000-C000-000000000046}  

A collection of all the ColorStop objects for the specified series.

**Remarks:** Each ColorStop object represents a color stop for a gradient fill in a range or selection.

**Example:**

```vba
With Selection.Interior
 .Pattern = xlPatternLinearGradient
 .Gradient.Degree = 90
 .Gradient.ColorStops.Clear
End With

 'adds stops after any have been cleared
With Selection.Interior.Gradient.ColorStops.Add(0)
 .ThemeColor = xlThemeColorDark1
 .TintAndShade = 0
End With

With Selection.Interior.Gradient.ColorStops.Add(1)
 .ThemeColor = xlThemeColorAccent1
 .TintAndShade = 0
End With
```

## Properties (6)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Count As Long  (read-only)`  
  Returns or sets the count of the represented object. Read-only.
- `_Default As ColorStop  (read-only)`
- `_NewEnum As IUnknown  (read-only)`

## Methods (3)

- `Add(Position As Double) As ColorStop`  
  Adds a ColorStop object to the specified collection.
    - `Position As Double` (required): Represents the position in which to apply the ColorStop.
- `Item(Index As Variant) As ColorStop`  
  Returns a single object from the represented collection.
    - `Index As Variant` (required): The name or index number for the object.
- `Clear()`  
  Clears the represented object.
