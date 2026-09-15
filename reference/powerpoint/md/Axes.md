# Axes

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {92D41A52-F07E-4CA4-AF6F-BEF486AA4E6F}  

Represents a collection of all the Axis objects in the specified chart.

**Remarks:** Use the Axes method to return the Axes collection. Use Axes ( _Type_, _AxisGroup_ ), where _Type_ is the axis type and _AxisGroup_ is the axis group, to return an Axes collection that contains a single Axis object. _Type_ can be one of the following XlAxisType constants: xlCategory, xlSeries, or xlValue. _AxisGroup_ can be one of the following XlAxisGroup constants: xlPrimary or xlSecondary.

**Example:**

```vba
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        MsgBox .Chart.Axes.Count

    End If

End With
```

## Properties (4)

- `Count As Long  (read-only)`  
  Returns the number of objects in the collection. Read-only Long.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent for the specified object. Read-only Object.
- `Application As Application  (read-only)`  
  When used without an object qualifier, returns an Application object that represents the Microsoft PowerPoint application. When used with an object qualifier, returns an Application object that represents the creator of the specified object (you can use this property with an Automation object to return the application of that object). Read-only.

## Methods (3)

- `Item(Type As XlAxisType, [AxisGroup As XlAxisGroup]) As Axis`  
  Returns a single Axis object from an Axes collection.
    - `Type As XlAxisType` (required): The axis type.
    - `AxisGroup As XlAxisGroup` (optional): The axis.
- `_NewEnum() As IUnknown`
- `_Default(Type As XlAxisType, [AxisGroup As XlAxisGroup]) As Axis`
