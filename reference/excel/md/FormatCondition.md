# FormatCondition

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024425-0000-0000-C000-000000000046}  

Represents a conditional format.

**Remarks:** The FormatCondition object is a member of the FormatConditions collection. The FormatConditions collection can now contain more than three conditional formats for a given range. Use the Add method of the FormatConditions object to create a new conditional format. If a range has multiple formats, you can use the Modify method to change one of the formats, or you can use the Delete method to delete a format, and then use the Add method to create a new format. Use the Font, Borders, and Interior properties of the FormatCondition object to control the appearance of formatted cells. Some properties of these objects aren't supported by the conditional format object model. Some of the properties that can be used with conditional formatting are listed in the following table.

**Example:**

```vba
With Worksheets(1).Range("e1:e10").FormatConditions(1)
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

## Properties (19)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Type As Long  (read-only)`  
  Returns a Long value, containing an XlFormatConditionType value, that represents the object type.
- `Operator As Long  (read-only)`  
  Returns a Long value that represents the operator for the conditional format.
- `Formula1 As String  (read-only)`  
  Returns the value or expression associated with the conditional format or data validation. Can be a constant value, a string value, a cell reference, or a formula. Read-only String.
- `Formula2 As String  (read-only)`  
  Returns the value or expression associated with the second part of a conditional format or data validation. Used only when the data validation conditional format Operator property is xlBetween or xlNotBetween. Can be a constant value, a string value, a cell reference, or a formula. Read-only String.
- `Interior As Interior  (read-only)`  
  Returns an Interior object that represents the interior of the specified object.
- `Borders As Borders  (read-only)`  
  Returns a Borders collection that represents the borders of a style or a range of cells (including a range defined as part of a conditional format).
- `Font As Font  (read-only)`  
  Returns a Font object that represents the font of the specified object.
- `Text As String  (read/write)`  
  Returns or sets a String value specifying the text string used by the conditional formatting rule.
- `TextOperator As XlContainsOperator  (read/write)`  
  Returns or sets one of the constants of the XlContainsOperator enumeration, specifying the text search performed by the conditional formatting rule.
- `DateOperator As XlTimePeriods  (read/write)`  
  Specifies the Date operator used in the format condition. Read/write.
- `NumberFormat As Variant  (read/write)`  
  Returns or sets the number format applied to a cell if the conditional formatting rule evaluates to True. Read/write Variant.
- `Priority As Long  (read/write)`  
  Returns or sets the priority value of the conditional formatting rule. The priority determines the order of evaluation when multiple conditional formatting rules exist on a worksheet.
- `StopIfTrue As Boolean  (read/write)`  
  Returns or sets a Boolean value that determines if additional formatting rules on the cell should be evaluated if the current rule evaluates to True.
- `AppliesTo As Range  (read-only)`  
  Returns a Range object specifying the cell range to which the formatting rule is applied.
- `PTCondition As Boolean  (read-only)`  
  Returns a Boolean value indicating if the conditional format is being applied to a PivotTable. Read-only.
- `ScopeType As XlPivotConditionScope  (read/write)`  
  Returns or sets one of the constants of the XlPivotConditionScope enumeration, which determines the scope of the conditional format when it is applied to a PivotTable.

## Methods (5)

- `Delete()`  
  Deletes the object.
- `Modify(Type As XlFormatConditionType, [Operator As Variant], [Formula1 As Variant], [Formula2 As Variant], [String As Variant], [Operator2 As Variant])`  
  Modifies an existing conditional format.
    - `Type As XlFormatConditionType` (required): Specifies whether the conditional format is based on a cell value or an expression.
    - `Operator As Variant` (optional): An XlFormatConditionOperator value that represents the conditional format operator. This parameter is ignored if _Type_ is set to xlExpression.
    - `Formula1 As Variant` (optional): The value or expression associated with the conditional format. Can be a constant value, a string value, a cell reference, or a formula.
    - `Formula2 As Variant` (optional): The value or expression associated with the conditional format. Can be a constant value, a string value, a cell reference, or a formula.
- `ModifyAppliesToRange(Range As Range)`  
  Sets the cell range to which this formatting rule applies.
    - `Range As Range` (required): The range to which this formatting rule will be applied.
- `SetFirstPriority()`  
  Sets the priority value for this conditional formatting rule to 1 so that it will be evaluated before all other rules on the worksheet.
- `SetLastPriority()`  
  Sets the evaluation order for this conditional formatting rule so that it is evaluated after all other rules on the worksheet.
