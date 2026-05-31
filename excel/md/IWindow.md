# IWindow

**Type:** Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00020893-0001-0000-C000-000000000046}  

## Properties (55)

- `Application As HRESULT  (read-only)`
- `Creator As HRESULT  (read-only)`
- `Parent As HRESULT  (read-only)`
- `ActiveCell As HRESULT  (read-only)`
- `ActiveChart As HRESULT  (read-only)`
- `ActivePane As HRESULT  (read-only)`
- `ActiveSheet As HRESULT  (read-only)`
- `Caption As HRESULT  (read/write)`
- `DisplayFormulas As HRESULT  (read/write)`
- `DisplayGridlines As HRESULT  (read/write)`
- `DisplayHeadings As HRESULT  (read/write)`
- `DisplayHorizontalScrollBar As HRESULT  (read/write)`
- `DisplayOutline As HRESULT  (read/write)`
- `DisplayVerticalScrollBar As HRESULT  (read/write)`
- `DisplayWorkbookTabs As HRESULT  (read/write)`
- `DisplayZeros As HRESULT  (read/write)`
- `EnableResize As HRESULT  (read/write)`
- `FreezePanes As HRESULT  (read/write)`
- `GridlineColor As HRESULT  (read/write)`
- `GridlineColorIndex As HRESULT  (read/write)`
- `Height As HRESULT  (read/write)`
- `Index As HRESULT  (read-only)`
- `Left As HRESULT  (read/write)`
- `OnWindow As HRESULT  (read/write)`
- `Panes As HRESULT  (read-only)`
- `RangeSelection As HRESULT  (read-only)`
- `ScrollColumn As HRESULT  (read/write)`
- `ScrollRow As HRESULT  (read/write)`
- `SelectedSheets As HRESULT  (read-only)`
- `Selection As HRESULT  (read-only)`
- `Split As HRESULT  (read/write)`
- `SplitColumn As HRESULT  (read/write)`
- `SplitHorizontal As HRESULT  (read/write)`
- `SplitRow As HRESULT  (read/write)`
- `SplitVertical As HRESULT  (read/write)`
- `TabRatio As HRESULT  (read/write)`
- `Top As HRESULT  (read/write)`
- `Type As HRESULT  (read-only)`
- `UsableHeight As HRESULT  (read-only)`
- `UsableWidth As HRESULT  (read-only)`
- `Visible As HRESULT  (read/write)`
- `VisibleRange As HRESULT  (read-only)`
- `Width As HRESULT  (read/write)`
- `WindowNumber As HRESULT  (read-only)`
- `WindowState As HRESULT  (read/write)`
- `Zoom As HRESULT  (read/write)`
- `View As HRESULT  (read/write)`
- `DisplayRightToLeft As HRESULT  (read/write)`
- `SheetViews As HRESULT  (read-only)`
- `ActiveSheetView As HRESULT  (read-only)`
- `DisplayRuler As HRESULT  (read/write)`
- `AutoFilterDateGrouping As HRESULT  (read/write)`
- `DisplayWhitespace As HRESULT  (read/write)`
- `Hwnd As HRESULT  (read-only)`
- `DisplayDataTypeIcons As HRESULT  (read/write)`

## Methods (14)

- `Activate(RHS As Variant)`
- `ActivateNext(RHS As Variant)`
- `ActivatePrevious(RHS As Variant)`
- `Close([SaveChanges As Variant], [Filename As Variant], [RouteWorkbook As Variant], RHS As Boolean)`
- `LargeScroll([Down As Variant], [Up As Variant], [ToRight As Variant], [ToLeft As Variant], RHS As Variant)`
- `NewWindow(RHS As Window)`
- `PrintPreview([EnableChanges As Variant], RHS As Variant)`
- `ScrollWorkbookTabs([Sheets As Variant], [Position As Variant], RHS As Variant)`
- `SmallScroll([Down As Variant], [Up As Variant], [ToRight As Variant], [ToLeft As Variant], RHS As Variant)`
- `PointsToScreenPixelsX(Points As Long, RHS As Long)`
- `PointsToScreenPixelsY(Points As Long, RHS As Long)`
- `RangeFromPoint(x As Long, y As Long, RHS As Object)`
- `ScrollIntoView(Left As Long, Top As Long, Width As Long, Height As Long, [Start As Variant])`
- `PrintOut([From As Variant], [To As Variant], [Copies As Variant], [Preview As Variant], [ActivePrinter As Variant], [PrintToFile As Variant], [Collate As Variant], [PrToFileName As Variant], RHS As Variant)`
