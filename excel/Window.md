# Window

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00020893-0000-0000-C000-000000000046}  

## Properties (55)

- `Application As Application  (read-only)`
- `Creator As XlCreator  (read-only)`
- `Parent As Object  (read-only)`
- `ActiveCell As Range  (read-only)`
- `ActiveChart As Chart  (read-only)`
- `ActivePane As Pane  (read-only)`
- `ActiveSheet As Object  (read-only)`
- `Caption As Variant  (read/write)`
- `DisplayFormulas As Boolean  (read/write)`
- `DisplayGridlines As Boolean  (read/write)`
- `DisplayHeadings As Boolean  (read/write)`
- `DisplayHorizontalScrollBar As Boolean  (read/write)`
- `DisplayOutline As Boolean  (read/write)`
- `DisplayVerticalScrollBar As Boolean  (read/write)`
- `DisplayWorkbookTabs As Boolean  (read/write)`
- `DisplayZeros As Boolean  (read/write)`
- `EnableResize As Boolean  (read/write)`
- `FreezePanes As Boolean  (read/write)`
- `GridlineColor As Long  (read/write)`
- `GridlineColorIndex As XlColorIndex  (read/write)`
- `Height As Double  (read/write)`
- `Index As Long  (read-only)`
- `Left As Double  (read/write)`
- `OnWindow As String  (read/write)`
- `Panes As Panes  (read-only)`
- `RangeSelection As Range  (read-only)`
- `ScrollColumn As Long  (read/write)`
- `ScrollRow As Long  (read/write)`
- `SelectedSheets As Sheets  (read-only)`
- `Selection As Object  (read-only)`
- `Split As Boolean  (read/write)`
- `SplitColumn As Long  (read/write)`
- `SplitHorizontal As Double  (read/write)`
- `SplitRow As Long  (read/write)`
- `SplitVertical As Double  (read/write)`
- `TabRatio As Double  (read/write)`
- `Top As Double  (read/write)`
- `Type As XlWindowType  (read-only)`
- `UsableHeight As Double  (read-only)`
- `UsableWidth As Double  (read-only)`
- `Visible As Boolean  (read/write)`
- `VisibleRange As Range  (read-only)`
- `Width As Double  (read/write)`
- `WindowNumber As Long  (read-only)`
- `WindowState As XlWindowState  (read/write)`
- `Zoom As Variant  (read/write)`
- `View As XlWindowView  (read/write)`
- `DisplayRightToLeft As Boolean  (read/write)`
- `SheetViews As SheetViews  (read-only)`
- `ActiveSheetView As Object  (read-only)`
- `DisplayRuler As Boolean  (read/write)`
- `AutoFilterDateGrouping As Boolean  (read/write)`
- `DisplayWhitespace As Boolean  (read/write)`
- `Hwnd As Long  (read-only)`
- `DisplayDataTypeIcons As Boolean  (read/write)`

## Methods (14)

- `Activate() As Variant`
- `ActivateNext() As Variant`
- `ActivatePrevious() As Variant`
- `Close([SaveChanges As Variant], [Filename As Variant], [RouteWorkbook As Variant]) As Boolean`
- `LargeScroll([Down As Variant], [Up As Variant], [ToRight As Variant], [ToLeft As Variant]) As Variant`
- `NewWindow() As Window`
- `PrintPreview([EnableChanges As Variant]) As Variant`
- `ScrollWorkbookTabs([Sheets As Variant], [Position As Variant]) As Variant`
- `SmallScroll([Down As Variant], [Up As Variant], [ToRight As Variant], [ToLeft As Variant]) As Variant`
- `PointsToScreenPixelsX(Points As Long) As Long`
- `PointsToScreenPixelsY(Points As Long) As Long`
- `RangeFromPoint(x As Long, y As Long) As Object`
- `ScrollIntoView(Left As Long, Top As Long, Width As Long, Height As Long, [Start As Variant])`
- `PrintOut([From As Variant], [To As Variant], [Copies As Variant], [Preview As Variant], [ActivePrinter As Variant], [PrintToFile As Variant], [Collate As Variant], [PrToFileName As Variant]) As Variant`
