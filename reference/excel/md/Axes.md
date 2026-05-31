# Axes

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {0002085B-0000-0000-C000-000000000046}  

A collection of all the Axis objects in the specified chart.

**Remarks:** Use the Axes method of the Chart object to return the Axes collection. Use Axes (_type_, _group_), where _type_ is the axis type and _group_ is the axis group, to return a single Axis object. - _Type_ can be one of the following XlAxisType constants: xlCategory, xlSeriesAxis, or xlValue. - _Group_ can be one of the following XlAxisGroup constants: xlPrimary or xlSecondary.

**Example:**

```vba
With Worksheets(1).ChartObjects(1).Chart
 MsgBox.Axes.Count
End With
```

## Properties (4)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Count As Long  (read-only)`  
  Returns a Long value that represents the number of objects in the collection.

## Methods (3)

- `Item(Type As XlAxisType, [AxisGroup As XlAxisGroup]) As Axis`  
  Returns a single Axis object from an Axes collection.
    - `Type As XlAxisType` (required): The axis type.
    - `AxisGroup As XlAxisGroup` (optional): The axis.
- `_NewEnum() As IUnknown`
- `_Default(Type As XlAxisType, [AxisGroup As XlAxisGroup]) As Axis`
