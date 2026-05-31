# DialogSheet

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000208AF-0000-0000-C000-000000000046}  

## Properties (42)

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
- `EnableCalculation As Boolean  (read/write)`
- `EnableAutoFilter As Boolean  (read/write)`
- `EnableSelection As XlEnableSelection  (read/write)`
- `EnableOutlining As Boolean  (read/write)`
- `EnablePivotTable As Boolean  (read/write)`
- `Names As Names  (read-only)`
- `ScrollArea As String  (read/write)`
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
- `EnableFormatConditionsCalculation As Boolean  (read/write)`
- `PrintedCommentPages As Long  (read-only)`
- `CommentsThreaded As CommentsThreaded  (read-only)`
- `AutoFilter As AutoFilter  (read-only)`
- `Sort As Sort  (read-only)`
- `NamedSheetViews As NamedSheetViewCollection  (read-only)`
- `DefaultButton As Variant  (read/write)`
- `Focus As Variant  (read/write)`

## Methods (23)

- `Activate()`
- `Copy([Before As Variant], [After As Variant])`
- `Delete()`
- `Move([Before As Variant], [After As Variant])`
- `PrintPreview([EnableChanges As Variant])`
- `Select([Replace As Variant])`
- `Unprotect([Password As Variant])`
- `ChartObjects([Index As Variant]) As Object`
- `CheckSpelling([CustomDictionary As Variant], [IgnoreUppercase As Variant], [AlwaysSuggest As Variant], [SpellLang As Variant])`
- `Evaluate(Name As Variant) As Variant`
- `_Evaluate(Name As Variant) As Variant`
- `ResetAllPageBreaks()`
- `OLEObjects([Index As Variant]) As Object`
- `Paste([Destination As Variant], [Link As Variant])`
- `ClearCircles()`
- `CircleInvalid()`
- `PasteSpecial([Format As Variant], [Link As Variant], [DisplayAsIcon As Variant], [IconFileName As Variant], [IconIndex As Variant], [IconLabel As Variant], [NoHTMLFormatting As Variant])`
- `Protect([Password As Variant], [DrawingObjects As Variant], [Contents As Variant], [Scenarios As Variant], [UserInterfaceOnly As Variant], [AllowFormattingCells As Variant], [AllowFormattingColumns As Variant], [AllowFormattingRows As Variant], [AllowInsertingColumns As Variant], [AllowInsertingRows As Variant], [AllowInsertingHyperlinks As Variant], [AllowDeletingColumns As Variant], [AllowDeletingRows As Variant], [AllowSorting As Variant], [AllowFiltering As Variant], [AllowUsingPivotTables As Variant])`
- `PrintOut([From As Variant], [To As Variant], [Copies As Variant], [Preview As Variant], [ActivePrinter As Variant], [PrintToFile As Variant], [Collate As Variant], [PrToFileName As Variant])`
- `ExportAsFixedFormat(Type As XlFixedFormatType, [Filename As Variant], [Quality As Variant], [IncludeDocProperties As Variant], [IgnorePrintAreas As Variant], [From As Variant], [To As Variant], [OpenAfterPublish As Variant], [FixedFormatExtClassPtr As Variant], [WorkIdentity As Variant])`
- `SaveAs(Filename As String, [FileFormat As Variant], [Password As Variant], [WriteResPassword As Variant], [ReadOnlyRecommended As Variant], [CreateBackup As Variant], [AddToMru As Variant], [TextCodepage As Variant], [TextVisualLayout As Variant], [Local As Variant])`
- `Hide([Cancel As Variant]) As Boolean`
- `Show() As Boolean`
