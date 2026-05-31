# FormatCondition

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024425-0000-0000-C000-000000000046}  

## Properties (19)

- `Application As Application  (read-only)`
- `Creator As XlCreator  (read-only)`
- `Parent As Object  (read-only)`
- `Type As Long  (read-only)`
- `Operator As Long  (read-only)`
- `Formula1 As String  (read-only)`
- `Formula2 As String  (read-only)`
- `Interior As Interior  (read-only)`
- `Borders As Borders  (read-only)`
- `Font As Font  (read-only)`
- `Text As String  (read/write)`
- `TextOperator As XlContainsOperator  (read/write)`
- `DateOperator As XlTimePeriods  (read/write)`
- `NumberFormat As Variant  (read/write)`
- `Priority As Long  (read/write)`
- `StopIfTrue As Boolean  (read/write)`
- `AppliesTo As Range  (read-only)`
- `PTCondition As Boolean  (read-only)`
- `ScopeType As XlPivotConditionScope  (read/write)`

## Methods (5)

- `Delete()`
- `Modify(Type As XlFormatConditionType, [Operator As Variant], [Formula1 As Variant], [Formula2 As Variant], [String As Variant], [Operator2 As Variant])`
- `ModifyAppliesToRange(Range As Range)`
- `SetFirstPriority()`
- `SetLastPriority()`
