# Worksheet

**Type:** Class  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00020820-0000-0000-C000-000000000046}  

## Properties (58)

- `Application As Application  (read-only)`
- `Creator As XlCreator  (read-only)`
- `Parent As Object  (read-only)`
- `CodeName As String  (read-only)`
- `_CodeName As String  (read/write)`
- `Index As Long  (read-only)`
- `Name As String  (read/write)`
- `Next As Object  (read-only)`
- `PageSetup As PageSetup  (read-only)`
- `Previous As Object  (read-only)`
- `ProtectContents As Boolean  (read-only)`
- `ProtectDrawingObjects As Boolean  (read-only)`
- `ProtectionMode As Boolean  (read-only)`
- `ProtectScenarios As Boolean  (read-only)`
- `Visible As XlSheetVisibility  (read/write)`
- `Shapes As Shapes  (read-only)`
- `TransitionExpEval As Boolean  (read/write)`
- `AutoFilterMode As Boolean  (read/write)`
- `EnableCalculation As Boolean  (read/write)`
- `Cells As Range  (read-only)`
- `CircularReference As Range  (read-only)`
- `Columns As Range  (read-only)`
- `ConsolidationFunction As XlConsolidationFunction  (read-only)`
- `ConsolidationOptions As Variant  (read-only)`
- `ConsolidationSources As Variant  (read-only)`
- `EnableAutoFilter As Boolean  (read/write)`
- `EnableSelection As XlEnableSelection  (read/write)`
- `EnableOutlining As Boolean  (read/write)`
- `EnablePivotTable As Boolean  (read/write)`
- `FilterMode As Boolean  (read-only)`
- `Names As Names  (read-only)`
- `Outline As Outline  (read-only)`
- `Range As Range  (read-only)`
- `Rows As Range  (read-only)`
- `ScrollArea As String  (read/write)`
- `StandardHeight As Double  (read-only)`
- `StandardWidth As Double  (read/write)`
- `TransitionFormEntry As Boolean  (read/write)`
- `Type As XlSheetType  (read-only)`
- `UsedRange As Range  (read-only)`
- `HPageBreaks As HPageBreaks  (read-only)`
- `VPageBreaks As VPageBreaks  (read-only)`
- `QueryTables As QueryTables  (read-only)`
- `DisplayPageBreaks As Boolean  (read/write)`
- `Comments As Comments  (read-only)`
- `Hyperlinks As Hyperlinks  (read-only)`
- `DisplayRightToLeft As Boolean  (read/write)`
- `Tab As Tab  (read-only)`
- `MailEnvelope As MsoEnvelope  (read-only)`
- `CustomProperties As CustomProperties  (read-only)`
- `Protection As Protection  (read-only)`
- `ListObjects As ListObjects  (read-only)`
- `EnableFormatConditionsCalculation As Boolean  (read/write)`
- `PrintedCommentPages As Long  (read-only)`
- `CommentsThreaded As CommentsThreaded  (read-only)`
- `AutoFilter As AutoFilter  (read-only)`
- `Sort As Sort  (read-only)`
- `NamedSheetViews As NamedSheetViewCollection  (read-only)`

## Methods (31)

- `Activate()`
- `Copy([Before As Variant], [After As Variant])`
- `Delete()`
- `Move([Before As Variant], [After As Variant])`
- `PrintPreview([EnableChanges As Variant])`
- `Select([Replace As Variant])`
- `Unprotect([Password As Variant])`
- `SetBackgroundPicture(Filename As String)`
- `Calculate()`
- `ChartObjects([Index As Variant]) As Object`
- `CheckSpelling([CustomDictionary As Variant], [IgnoreUppercase As Variant], [AlwaysSuggest As Variant], [SpellLang As Variant])`
- `ClearArrows()`
- `Evaluate(Name As Variant) As Variant`
- `_Evaluate(Name As Variant) As Variant`
- `ResetAllPageBreaks()`
- `OLEObjects([Index As Variant]) As Object`
- `Paste([Destination As Variant], [Link As Variant])`
- `PivotTables([Index As Variant]) As Object`
- `PivotTableWizard([SourceType As Variant], [SourceData As Variant], [TableDestination As Variant], [TableName As Variant], [RowGrand As Variant], [ColumnGrand As Variant], [SaveData As Variant], [HasAutoFormat As Variant], [AutoPage As Variant], [Reserved As Variant], [BackgroundQuery As Variant], [OptimizeCache As Variant], [PageFieldOrder As Variant], [PageFieldWrapCount As Variant], [ReadData As Variant], [Connection As Variant]) As PivotTable`
- `Scenarios([Index As Variant]) As Object`
- `ShowAllData()`
- `ShowDataForm()`
- `ClearCircles()`
- `CircleInvalid()`
- `PasteSpecial([Format As Variant], [Link As Variant], [DisplayAsIcon As Variant], [IconFileName As Variant], [IconIndex As Variant], [IconLabel As Variant], [NoHTMLFormatting As Variant])`
- `Protect([Password As Variant], [DrawingObjects As Variant], [Contents As Variant], [Scenarios As Variant], [UserInterfaceOnly As Variant], [AllowFormattingCells As Variant], [AllowFormattingColumns As Variant], [AllowFormattingRows As Variant], [AllowInsertingColumns As Variant], [AllowInsertingRows As Variant], [AllowInsertingHyperlinks As Variant], [AllowDeletingColumns As Variant], [AllowDeletingRows As Variant], [AllowSorting As Variant], [AllowFiltering As Variant], [AllowUsingPivotTables As Variant])`
- `XmlDataQuery(XPath As String, [SelectionNamespaces As Variant], [Map As Variant]) As Range`
- `XmlMapQuery(XPath As String, [SelectionNamespaces As Variant], [Map As Variant]) As Range`
- `PrintOut([From As Variant], [To As Variant], [Copies As Variant], [Preview As Variant], [ActivePrinter As Variant], [PrintToFile As Variant], [Collate As Variant], [PrToFileName As Variant], [IgnorePrintAreas As Variant])`
- `ExportAsFixedFormat(Type As XlFixedFormatType, [Filename As Variant], [Quality As Variant], [IncludeDocProperties As Variant], [IgnorePrintAreas As Variant], [From As Variant], [To As Variant], [OpenAfterPublish As Variant], [FixedFormatExtClassPtr As Variant], [WorkIdentity As Variant])`
- `SaveAs(Filename As String, [FileFormat As Variant], [Password As Variant], [WriteResPassword As Variant], [ReadOnlyRecommended As Variant], [CreateBackup As Variant], [AddToMru As Variant], [TextCodepage As Variant], [TextVisualLayout As Variant], [Local As Variant])`

## Events (17)

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
