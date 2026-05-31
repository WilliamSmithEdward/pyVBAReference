# IScenarios

**Type:** Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00020896-0001-0000-C000-000000000046}  

## Properties (4)

- `Application As HRESULT  (read-only)`
- `Creator As HRESULT  (read-only)`
- `Parent As HRESULT  (read-only)`
- `Count As HRESULT  (read-only)`

## Methods (5)

- `Add(Name As String, ChangingCells As Variant, [Values As Variant], [Comment As Variant], [Locked As Variant], [Hidden As Variant], RHS As Scenario)`
- `CreateSummary([ReportType As XlSummaryReportType], [ResultCells As Variant], RHS As Variant)`
- `Item(Index As Variant, RHS As Scenario)`
- `Merge(Source As Variant, RHS As Variant)`
- `_NewEnum(RHS As IUnknown)`
