# IAppEvents

**Type:** Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024413-0001-0000-C000-000000000046}  

## Methods (49)

- `NewWorkbook(Wb As Workbook)`
- `SheetSelectionChange(Sh As Object, Target As Range)`
- `SheetBeforeDoubleClick(Sh As Object, Target As Range, Cancel As Boolean)`
- `SheetBeforeRightClick(Sh As Object, Target As Range, Cancel As Boolean)`
- `SheetActivate(Sh As Object)`
- `SheetDeactivate(Sh As Object)`
- `SheetCalculate(Sh As Object)`
- `SheetChange(Sh As Object, Target As Range)`
- `WorkbookOpen(Wb As Workbook)`
- `WorkbookActivate(Wb As Workbook)`
- `WorkbookDeactivate(Wb As Workbook)`
- `WorkbookBeforeClose(Wb As Workbook, Cancel As Boolean)`
- `WorkbookBeforeSave(Wb As Workbook, SaveAsUI As Boolean, Cancel As Boolean)`
- `WorkbookBeforePrint(Wb As Workbook, Cancel As Boolean)`
- `WorkbookNewSheet(Wb As Workbook, Sh As Object)`
- `WorkbookAddinInstall(Wb As Workbook)`
- `WorkbookAddinUninstall(Wb As Workbook)`
- `WindowResize(Wb As Workbook, Wn As Window)`
- `WindowActivate(Wb As Workbook, Wn As Window)`
- `WindowDeactivate(Wb As Workbook, Wn As Window)`
- `SheetFollowHyperlink(Sh As Object, Target As Hyperlink)`
- `SheetPivotTableUpdate(Sh As Object, Target As PivotTable)`
- `WorkbookPivotTableCloseConnection(Wb As Workbook, Target As PivotTable)`
- `WorkbookPivotTableOpenConnection(Wb As Workbook, Target As PivotTable)`
- `WorkbookSync(Wb As Workbook, SyncEventType As MsoSyncEventType)`
- `WorkbookBeforeXmlImport(Wb As Workbook, Map As XmlMap, Url As String, IsRefresh As Boolean, Cancel As Boolean)`
- `WorkbookAfterXmlImport(Wb As Workbook, Map As XmlMap, IsRefresh As Boolean, Result As XlXmlImportResult)`
- `WorkbookBeforeXmlExport(Wb As Workbook, Map As XmlMap, Url As String, Cancel As Boolean)`
- `WorkbookAfterXmlExport(Wb As Workbook, Map As XmlMap, Url As String, Result As XlXmlExportResult)`
- `WorkbookRowsetComplete(Wb As Workbook, Description As String, Sheet As String, Success As Boolean)`
- `AfterCalculate()`
- `SheetPivotTableAfterValueChange(Sh As Object, TargetPivotTable As PivotTable, TargetRange As Range)`
- `SheetPivotTableBeforeAllocateChanges(Sh As Object, TargetPivotTable As PivotTable, ValueChangeStart As Long, ValueChangeEnd As Long, Cancel As Boolean)`
- `SheetPivotTableBeforeCommitChanges(Sh As Object, TargetPivotTable As PivotTable, ValueChangeStart As Long, ValueChangeEnd As Long, Cancel As Boolean)`
- `SheetPivotTableBeforeDiscardChanges(Sh As Object, TargetPivotTable As PivotTable, ValueChangeStart As Long, ValueChangeEnd As Long)`
- `ProtectedViewWindowOpen(Pvw As ProtectedViewWindow)`
- `ProtectedViewWindowBeforeEdit(Pvw As ProtectedViewWindow, Cancel As Boolean)`
- `ProtectedViewWindowBeforeClose(Pvw As ProtectedViewWindow, Reason As XlProtectedViewCloseReason, Cancel As Boolean)`
- `ProtectedViewWindowResize(Pvw As ProtectedViewWindow)`
- `ProtectedViewWindowActivate(Pvw As ProtectedViewWindow)`
- `ProtectedViewWindowDeactivate(Pvw As ProtectedViewWindow)`
- `WorkbookAfterSave(Wb As Workbook, Success As Boolean)`
- `WorkbookNewChart(Wb As Workbook, Ch As Chart)`
- `SheetLensGalleryRenderComplete(Sh As Object)`
- `SheetTableUpdate(Sh As Object, Target As TableObject)`
- `WorkbookModelChange(Wb As Workbook, Changes As ModelChanges)`
- `SheetBeforeDelete(Sh As Object)`
- `WorkbookBeforeRemoteChange(Wb As Workbook)`
- `WorkbookAfterRemoteChange(Wb As Workbook)`
