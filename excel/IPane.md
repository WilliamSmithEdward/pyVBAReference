# IPane

**Type:** Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00020895-0001-0000-C000-000000000046}  

## Properties (7)

- `Application As HRESULT  (read-only)`
- `Creator As HRESULT  (read-only)`
- `Parent As HRESULT  (read-only)`
- `Index As HRESULT  (read-only)`
- `ScrollColumn As HRESULT  (read/write)`
- `ScrollRow As HRESULT  (read/write)`
- `VisibleRange As HRESULT  (read-only)`

## Methods (6)

- `Activate(RHS As Boolean)`
- `LargeScroll([Down As Variant], [Up As Variant], [ToRight As Variant], [ToLeft As Variant], RHS As Variant)`
- `SmallScroll([Down As Variant], [Up As Variant], [ToRight As Variant], [ToLeft As Variant], RHS As Variant)`
- `ScrollIntoView(Left As Long, Top As Long, Width As Long, Height As Long, [Start As Variant])`
- `PointsToScreenPixelsX(Points As Long, RHS As Long)`
- `PointsToScreenPixelsY(Points As Long, RHS As Long)`
