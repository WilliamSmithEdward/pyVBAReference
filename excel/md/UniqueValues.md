# UniqueValues

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {0002449F-0000-0000-C000-000000000046}  

The UniqueValues object uses the DupeUnique property to return or set an enum that determines whether the rule should look for duplicate or unique values in the range.

## Properties (14)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Priority As Long  (read/write)`  
  Returns or sets the priority value of the conditional formatting rule. The priority determines the order of evaluation when multiple conditional formatting rules exist on a worksheet.
- `StopIfTrue As Boolean  (read/write)`  
  Returns or sets a Boolean value that determines if additional formatting rules on the cell should be evaluated if the current rule evaluates to True.
- `AppliesTo As Range  (read-only)`  
  Returns a Range object specifying the cell range to which the formatting rule is applied.
- `DupeUnique As XlDupeUnique  (read/write)`  
  Returns or sets one of the constants of the XlDupeUnique enumeration, which specifies if the conditional formatting rule is looking for unique or duplicate values.
- `Interior As Interior  (read-only)`  
  Returns an Interior object that specifies a cell's interior attributes for a conditional formatting rule that evaluates to True. Read-only.
- `Borders As Borders  (read-only)`  
  Returns a Borders collection that specifies the formatting of cell borders if the conditional formatting rule evaluates to True. Read-only.
- `Font As Font  (read-only)`  
  Returns a Font object that specifies the font formatting if the conditional formatting rule evaluates to True. Read-only.
- `Type As Long  (read-only)`  
  Returns one of the constants of the XlFormatConditionType enumeration, which specifies the type of conditional format. Read-only.
- `NumberFormat As Variant  (read/write)`  
  Returns or sets the number format applied to a cell if the conditional formatting rule evaluates to True. Read/write Variant.
- `PTCondition As Boolean  (read-only)`  
  Returns a Boolean value indicating if the conditional format is being applied to a PivotTable. Read-only.
- `ScopeType As XlPivotConditionScope  (read/write)`  
  Returns or sets one of the constants of the XlPivotConditionScope enumeration, which determines the scope of the conditional format when it is applied to a PivotTable.

## Methods (4)

- `SetFirstPriority()`  
  Sets the priority value for this conditional formatting rule to 1 so that it will be evaluated before all other rules on the worksheet.
- `SetLastPriority()`  
  Sets the evaluation order for this conditional formatting rule so that it is evaluated after all other rules on the worksheet.
- `Delete()`  
  Deletes the specified conditional formatting rule object.
- `ModifyAppliesToRange(Range As Range)`  
  Sets the cell range to which this formatting rule applies.
    - `Range As Range` (required): The range to which this formatting rule will be applied.
