# FormatConditions

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024424-0000-0000-C000-000000000046}  

## Properties (6)

- `Application As Application  (read-only)`
- `Creator As XlCreator  (read-only)`
- `Parent As Object  (read-only)`
- `Count As Long  (read-only)`
- `_Default As Object  (read-only)`
- `_NewEnum As IUnknown  (read-only)`

## Methods (9)

- `Item(Index As Variant) As Object`
- `Add(Type As XlFormatConditionType, [Operator As Variant], [Formula1 As Variant], [Formula2 As Variant], [String As Variant], [TextOperator As Variant], [DateOperator As Variant], [ScopeType As Variant]) As Object`
- `Delete()`
- `AddColorScale(ColorScaleType As Long) As Object`
- `AddDatabar() As Object`
- `AddIconSetCondition() As Object`
- `AddTop10() As Object`
- `AddAboveAverage() As Object`
- `AddUniqueValues() As Object`
