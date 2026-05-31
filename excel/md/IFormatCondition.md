# IFormatCondition

**Type:** Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024425-0001-0000-C000-000000000046}  

## Properties (19)

- `Application As HRESULT  (read-only)`
- `Creator As HRESULT  (read-only)`
- `Parent As HRESULT  (read-only)`
- `Type As HRESULT  (read-only)`
- `Operator As HRESULT  (read-only)`
- `Formula1 As HRESULT  (read-only)`
- `Formula2 As HRESULT  (read-only)`
- `Interior As HRESULT  (read-only)`
- `Borders As HRESULT  (read-only)`
- `Font As HRESULT  (read-only)`
- `Text As HRESULT  (read/write)`
- `TextOperator As HRESULT  (read/write)`
- `DateOperator As HRESULT  (read/write)`
- `NumberFormat As HRESULT  (read/write)`
- `Priority As HRESULT  (read/write)`
- `StopIfTrue As HRESULT  (read/write)`
- `AppliesTo As HRESULT  (read-only)`
- `PTCondition As HRESULT  (read-only)`
- `ScopeType As HRESULT  (read/write)`

## Methods (5)

- `Delete()`
- `Modify(Type As XlFormatConditionType, [Operator As Variant], [Formula1 As Variant], [Formula2 As Variant], [String As Variant], [Operator2 As Variant])`
- `ModifyAppliesToRange(Range As Range)`
- `SetFirstPriority()`
- `SetLastPriority()`
