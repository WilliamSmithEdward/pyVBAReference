# WorkbookEvents

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024412-0000-0000-C000-000000000046}  

## Methods (42)

- `Open()`
- `Activate()`
- `Deactivate()`
- `BeforeClose(Cancel As Boolean)`
- `BeforeSave(SaveAsUI As Boolean, Cancel As Boolean)`
- `BeforePrint(Cancel As Boolean)`
- `NewSheet(Sh As Object)`
- `AddinInstall()`
- `AddinUninstall()`
- `WindowResize(Wn As Window)`
- `WindowActivate(Wn As Window)`
- `WindowDeactivate(Wn As Window)`
- `SheetSelectionChange(Sh As Object, Target As Range)`
- `SheetBeforeDoubleClick(Sh As Object, Target As Range, Cancel As Boolean)`
- `SheetBeforeRightClick(Sh As Object, Target As Range, Cancel As Boolean)`
- `SheetActivate(Sh As Object)`
- `SheetDeactivate(Sh As Object)`
- `SheetCalculate(Sh As Object)`
- `SheetChange(Sh As Object, Target As Range)`
- `SheetFollowHyperlink(Sh As Object, Target As Hyperlink)`
- `SheetPivotTableUpdate(Sh As Object, Target As PivotTable)`
- `PivotTableCloseConnection(Target As PivotTable)`
- `PivotTableOpenConnection(Target As PivotTable)`
- `Sync(SyncEventType As MsoSyncEventType)`
- `BeforeXmlImport(Map As XmlMap, Url As String, IsRefresh As Boolean, Cancel As Boolean)`
- `AfterXmlImport(Map As XmlMap, IsRefresh As Boolean, Result As XlXmlImportResult)`
- `BeforeXmlExport(Map As XmlMap, Url As String, Cancel As Boolean)`
- `AfterXmlExport(Map As XmlMap, Url As String, Result As XlXmlExportResult)`
- `RowsetComplete(Description As String, Sheet As String, Success As Boolean)`
- `SheetPivotTableAfterValueChange(Sh As Object, TargetPivotTable As PivotTable, TargetRange As Range)`
- `SheetPivotTableBeforeAllocateChanges(Sh As Object, TargetPivotTable As PivotTable, ValueChangeStart As Long, ValueChangeEnd As Long, Cancel As Boolean)`
- `SheetPivotTableBeforeCommitChanges(Sh As Object, TargetPivotTable As PivotTable, ValueChangeStart As Long, ValueChangeEnd As Long, Cancel As Boolean)`
- `SheetPivotTableBeforeDiscardChanges(Sh As Object, TargetPivotTable As PivotTable, ValueChangeStart As Long, ValueChangeEnd As Long)`
- `SheetPivotTableChangeSync(Sh As Object, Target As PivotTable)`
- `AfterSave(Success As Boolean)`
- `NewChart(Ch As Chart)`
- `SheetLensGalleryRenderComplete(Sh As Object)`
- `SheetTableUpdate(Sh As Object, Target As TableObject)`
- `ModelChange(Changes As ModelChanges)`
- `SheetBeforeDelete(Sh As Object)`
- `BeforeRemoteChange()`
- `AfterRemoteChange()`
