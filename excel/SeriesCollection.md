# SeriesCollection

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {0002086C-0000-0000-C000-000000000046}  

## Properties (4)

- `Application As Application  (read-only)`
- `Creator As XlCreator  (read-only)`
- `Parent As Object  (read-only)`
- `Count As Long  (read-only)`

## Methods (7)

- `Add(Source As Variant, [Rowcol As XlRowCol], [SeriesLabels As Variant], [CategoryLabels As Variant], [Replace As Variant]) As Series`
- `Extend(Source As Variant, [Rowcol As Variant], [CategoryLabels As Variant]) As Variant`
- `Item(Index As Variant) As Series`
- `_NewEnum() As IUnknown`
- `Paste([Rowcol As XlRowCol], [SeriesLabels As Variant], [CategoryLabels As Variant], [Replace As Variant], [NewSeries As Variant]) As Variant`
- `NewSeries() As Series`
- `_Default(Index As Variant) As Series`
