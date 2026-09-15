# Axes

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {354AB591-A217-48B4-99E4-14F58F15667D}  

Represents a collection of all the Axis objects in the specified chart.

**Remarks:** Use the Axes method to return the Axes collection. Use Axes ( _Type_ , _AxisGroup_ ), where _Type_ is the axis type and _AxisGroup_ is the axis group, to return an Axes collection that contains a single Axis object. _Type_ can be one of the following XlAxisType constants: xlCategory, xlSeries, or xlValue. _AxisGroup_ can be one of the following XlAxisGroup constants: xlPrimary or xlSecondary.

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
- `Application As Object  (read-only)`  
  When used without an object qualifier, returns an Application object that represents the Microsoft Word application. When used with an object qualifier, returns an Application object that represents the creator of the specified object (you can use this property with an Automation object to return the application of that object). Read-only.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent for the specified object. Read-only Object.

## Methods (3)

- `Item(Type As XlAxisType, [AxisGroup As XlAxisGroup]) As Axis`  
  Returns a single Axis object from an Axes collection.
    - `Type As XlAxisType` (required): One of the enumeration values that specifies the axis type.
    - `AxisGroup As XlAxisGroup` (optional): One of the enumeration values that specifies the axis.
- `_NewEnum() As IUnknown`
- `_Default(Type As XlAxisType, [AxisGroup As XlAxisGroup]) As Axis`
