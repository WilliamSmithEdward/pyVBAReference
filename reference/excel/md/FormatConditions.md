# FormatConditions

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024424-0000-0000-C000-000000000046}  

Represents the collection of conditional formats for a single range.

**Remarks:** The FormatConditions collection can contain multiple conditional formats. Each format is represented by a FormatCondition object. Use the FormatConditions property to return a FormatConditions object. Use the Add method to create a new conditional format, and use the Modify method of the FormatCondition object to change an existing conditional format.

**Example:**

```vba
With Worksheets(1).Range("e1:e10").FormatConditions _
 .Add(xlCellValue, xlGreater, "=$a$1")
 With .Borders
 .LineStyle = xlContinuous
 .Weight = xlThin
 .ColorIndex = 6
 End With
 With .Font
 .Bold = True
 .ColorIndex = 3
 End With
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
  Returns a Long value that represents the number of objects in the collection.
- `_Default As Object  (read-only)`
- `_NewEnum As IUnknown  (read-only)`

## Methods (9)

- `Item(Index As Variant) As Object`  
  Returns a single object from a collection.
    - `Index As Variant` (required): The name or index number for the object.
- `Add(Type As XlFormatConditionType, [Operator As Variant], [Formula1 As Variant], [Formula2 As Variant], [String As Variant], [TextOperator As Variant], [DateOperator As Variant], [ScopeType As Variant]) As Object`  
  Adds a new conditional format.
    - `Type As XlFormatConditionType` (required): Specifies whether the conditional format is based on a cell value or an expression.
    - `Operator As Variant` (optional): The conditional format operator. Can be one of the following XlFormatConditionOperator constants: xlBetween, xlEqual, xlGreater, xlGreaterEqual, xlLess, xlLessEqual, xlNotBetween, or xlNotEqual. If _Type_ is xlExpression, the _Operator_ argument is ignored.
    - `Formula1 As Variant` (optional): The value or expression associated with the conditional format. Can be a constant value, a string value, a cell reference, or a formula.
    - `Formula2 As Variant` (optional): The value or expression associated with the second part of the conditional format when _Operator_ is xlBetween or xlNotBetween (otherwise, this argument is ignored). Can be a constant value, a string value, a cell reference, or a formula.
- `Delete()`  
  Deletes the object.
- `AddColorScale(ColorScaleType As Long) As Object`  
  Returns a new ColorScale object representing a conditional formatting rule that uses gradations in cell colors to indicate relative differences in the values of cells included in a selected range.
    - `ColorScaleType As Long` (required): The type of color scale.
- `AddDatabar() As Object`  
  Returns a Databar object representing a data bar conditional formatting rule for the specified range.
- `AddIconSetCondition() As Object`  
  Returns a new IconSetCondition object that represents an icon set conditional formatting rule for the specified range.
- `AddTop10() As Object`  
  Returns a Top10 object representing a conditional formatting rule for the specified range.
- `AddAboveAverage() As Object`  
  Returns a new AboveAverage object representing a conditional formatting rule for the specified range.
- `AddUniqueValues() As Object`  
  Returns a new UniqueValues object representing a conditional formatting rule for the specified range.
