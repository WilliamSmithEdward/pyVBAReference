# IDialogSheet

**Type:** Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000208AF-0001-0000-C000-000000000046}  

## Properties (42)

- `Application As HRESULT  (read-only)`
- `Creator As HRESULT  (read-only)`
- `Parent As HRESULT  (read-only)`
- `CodeName As HRESULT  (read-only)`
- `_CodeName As HRESULT  (read/write)`
- `Index As HRESULT  (read-only)`
- `Name As HRESULT  (read/write)`
- `Next As HRESULT  (read-only)`
- `PageSetup As HRESULT  (read-only)`
- `Previous As HRESULT  (read-only)`
- `ProtectContents As HRESULT  (read-only)`
- `ProtectDrawingObjects As HRESULT  (read-only)`
- `ProtectionMode As HRESULT  (read-only)`
- `ProtectScenarios As HRESULT  (read-only)`
- `Visible As HRESULT  (read/write)`
- `Shapes As HRESULT  (read-only)`
- `EnableCalculation As HRESULT  (read/write)`
- `EnableAutoFilter As HRESULT  (read/write)`
- `EnableSelection As HRESULT  (read/write)`
- `EnableOutlining As HRESULT  (read/write)`
- `EnablePivotTable As HRESULT  (read/write)`
- `Names As HRESULT  (read-only)`
- `ScrollArea As HRESULT  (read/write)`
- `HPageBreaks As HRESULT  (read-only)`
- `VPageBreaks As HRESULT  (read-only)`
- `QueryTables As HRESULT  (read-only)`
- `DisplayPageBreaks As HRESULT  (read/write)`
- `Comments As HRESULT  (read-only)`
- `Hyperlinks As HRESULT  (read-only)`
- `DisplayRightToLeft As HRESULT  (read/write)`
- `Tab As HRESULT  (read-only)`
- `MailEnvelope As HRESULT  (read-only)`
- `CustomProperties As HRESULT  (read-only)`
- `Protection As HRESULT  (read-only)`
- `EnableFormatConditionsCalculation As HRESULT  (read/write)`
- `PrintedCommentPages As HRESULT  (read-only)`
- `CommentsThreaded As HRESULT  (read-only)`
- `AutoFilter As HRESULT  (read-only)`
- `Sort As HRESULT  (read-only)`
- `NamedSheetViews As HRESULT  (read-only)`
- `DefaultButton As HRESULT  (read/write)`
- `Focus As HRESULT  (read/write)`

## Methods (23)

- `Activate(lcid As Long)`
- `Copy([Before As Variant], [After As Variant], lcid As Long)`
- `Delete(lcid As Long)`
- `Move([Before As Variant], [After As Variant], lcid As Long)`
- `PrintPreview([EnableChanges As Variant], lcid As Long)`
- `Select([Replace As Variant], lcid As Long)`
- `Unprotect([Password As Variant], lcid As Long)`
- `ChartObjects([Index As Variant], lcid As Long, RHS As Object)`
- `CheckSpelling([CustomDictionary As Variant], [IgnoreUppercase As Variant], [AlwaysSuggest As Variant], [SpellLang As Variant], lcid As Long)`
- `Evaluate(Name As Variant, lcid As Long, RHS As Variant)`
- `_Evaluate(Name As Variant, lcid As Long, RHS As Variant)`
- `ResetAllPageBreaks()`
- `OLEObjects([Index As Variant], lcid As Long, RHS As Object)`
- `Paste([Destination As Variant], [Link As Variant], lcid As Long)`
- `ClearCircles()`
- `CircleInvalid()`
- `PasteSpecial([Format As Variant], [Link As Variant], [DisplayAsIcon As Variant], [IconFileName As Variant], [IconIndex As Variant], [IconLabel As Variant], [NoHTMLFormatting As Variant], lcid As Long)`
- `Protect([Password As Variant], [DrawingObjects As Variant], [Contents As Variant], [Scenarios As Variant], [UserInterfaceOnly As Variant], [AllowFormattingCells As Variant], [AllowFormattingColumns As Variant], [AllowFormattingRows As Variant], [AllowInsertingColumns As Variant], [AllowInsertingRows As Variant], [AllowInsertingHyperlinks As Variant], [AllowDeletingColumns As Variant], [AllowDeletingRows As Variant], [AllowSorting As Variant], [AllowFiltering As Variant], [AllowUsingPivotTables As Variant])`
- `PrintOut([From As Variant], [To As Variant], [Copies As Variant], [Preview As Variant], [ActivePrinter As Variant], [PrintToFile As Variant], [Collate As Variant], [PrToFileName As Variant])`
- `ExportAsFixedFormat(Type As XlFixedFormatType, [Filename As Variant], [Quality As Variant], [IncludeDocProperties As Variant], [IgnorePrintAreas As Variant], [From As Variant], [To As Variant], [OpenAfterPublish As Variant], [FixedFormatExtClassPtr As Variant], [WorkIdentity As Variant])`
- `SaveAs(Filename As String, [FileFormat As Variant], [Password As Variant], [WriteResPassword As Variant], [ReadOnlyRecommended As Variant], [CreateBackup As Variant], [AddToMru As Variant], [TextCodepage As Variant], [TextVisualLayout As Variant], [Local As Variant])`
- `Hide([Cancel As Variant], RHS As Boolean)`
- `Show(RHS As Boolean)`
