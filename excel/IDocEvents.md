# IDocEvents

**Type:** Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024411-0001-0000-C000-000000000046}  

## Methods (17)

- `SelectionChange(Target As Range)`
- `BeforeDoubleClick(Target As Range, Cancel As Boolean)`
- `BeforeRightClick(Target As Range, Cancel As Boolean)`
- `Activate()`
- `Deactivate()`
- `Calculate()`
- `Change(Target As Range)`
- `FollowHyperlink(Target As Hyperlink)`
- `PivotTableUpdate(Target As PivotTable)`
- `PivotTableAfterValueChange(TargetPivotTable As PivotTable, TargetRange As Range)`
- `PivotTableBeforeAllocateChanges(TargetPivotTable As PivotTable, ValueChangeStart As Long, ValueChangeEnd As Long, Cancel As Boolean)`
- `PivotTableBeforeCommitChanges(TargetPivotTable As PivotTable, ValueChangeStart As Long, ValueChangeEnd As Long, Cancel As Boolean)`
- `PivotTableBeforeDiscardChanges(TargetPivotTable As PivotTable, ValueChangeStart As Long, ValueChangeEnd As Long)`
- `PivotTableChangeSync(Target As PivotTable)`
- `LensGalleryRenderComplete()`
- `TableUpdate(Target As TableObject)`
- `BeforeDelete()`
