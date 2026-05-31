# IFormatConditions

**Type:** Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024424-0001-0000-C000-000000000046}  

## Properties (6)

- `Application As HRESULT  (read-only)`
- `Creator As HRESULT  (read-only)`
- `Parent As HRESULT  (read-only)`
- `Count As HRESULT  (read-only)`
- `_Default As HRESULT  (read-only)`
- `_NewEnum As HRESULT  (read-only)`

## Methods (9)

- `Item(Index As Variant, RHS As Object)`
- `Add(Type As XlFormatConditionType, [Operator As Variant], [Formula1 As Variant], [Formula2 As Variant], [String As Variant], [TextOperator As Variant], [DateOperator As Variant], [ScopeType As Variant], RHS As Object)`
- `Delete()`
- `AddColorScale(ColorScaleType As Long, RHS As Object)`
- `AddDatabar(RHS As Object)`
- `AddIconSetCondition(RHS As Object)`
- `AddTop10(RHS As Object)`
- `AddAboveAverage(RHS As Object)`
- `AddUniqueValues(RHS As Object)`
