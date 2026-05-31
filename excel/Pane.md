# Pane

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00020895-0000-0000-C000-000000000046}  

## Properties (7)

- `Application As Application  (read-only)`
- `Creator As XlCreator  (read-only)`
- `Parent As Object  (read-only)`
- `Index As Long  (read-only)`
- `ScrollColumn As Long  (read/write)`
- `ScrollRow As Long  (read/write)`
- `VisibleRange As Range  (read-only)`

## Methods (6)

- `Activate() As Boolean`
- `LargeScroll([Down As Variant], [Up As Variant], [ToRight As Variant], [ToLeft As Variant]) As Variant`
- `SmallScroll([Down As Variant], [Up As Variant], [ToRight As Variant], [ToLeft As Variant]) As Variant`
- `ScrollIntoView(Left As Long, Top As Long, Width As Long, Height As Long, [Start As Variant])`
- `PointsToScreenPixelsX(Points As Long) As Long`
- `PointsToScreenPixelsY(Points As Long) As Long`
