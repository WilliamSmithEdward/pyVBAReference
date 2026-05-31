# ISeriesCollection

**Type:** Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {0002086C-0001-0000-C000-000000000046}  

## Properties (4)

- `Application As HRESULT  (read-only)`
- `Creator As HRESULT  (read-only)`
- `Parent As HRESULT  (read-only)`
- `Count As HRESULT  (read-only)`

## Methods (7)

- `Add(Source As Variant, [Rowcol As XlRowCol], [SeriesLabels As Variant], [CategoryLabels As Variant], [Replace As Variant], RHS As Series)`
- `Extend(Source As Variant, [Rowcol As Variant], [CategoryLabels As Variant], RHS As Variant)`
- `Item(Index As Variant, RHS As Series)`
- `_NewEnum(RHS As IUnknown)`
- `Paste([Rowcol As XlRowCol], [SeriesLabels As Variant], [CategoryLabels As Variant], [Replace As Variant], [NewSeries As Variant], RHS As Variant)`
- `NewSeries(RHS As Series)`
- `_Default(Index As Variant, RHS As Series)`
