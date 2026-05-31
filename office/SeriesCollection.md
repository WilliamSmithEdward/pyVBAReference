# SeriesCollection

**Type:** Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C170A-0000-0000-C000-000000000046}  

## Properties (5)

- `Parent As HRESULT  (read-only)`
- `Count As HRESULT  (read-only)`
- `Application As HRESULT  (read-only)`
- `Creator As HRESULT  (read-only)`
- `_Default As HRESULT  (read-only)`

## Methods (6)

- `Add(Source As Variant, [Rowcol As XlRowCol], [SeriesLabels As Variant], [CategoryLabels As Variant], [Replace As Variant], RHS As IMsoSeries)`
- `Extend(Source As Variant, [Rowcol As Variant], [CategoryLabels As Variant], RHS As Variant)`
- `Item(Index As Variant, RHS As IMsoSeries)`
- `_NewEnum(RHS As IUnknown)`
- `Paste([Rowcol As XlRowCol], [SeriesLabels As Variant], [CategoryLabels As Variant], [Replace As Variant], [NewSeries As Variant], RHS As Variant)`
- `NewSeries(RHS As IMsoSeries)`
