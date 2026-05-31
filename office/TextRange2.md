# TextRange2

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C0397-0000-0000-C000-000000000046}  

## Properties (22)

- `Application As Object  (read-only)`
- `Creator As Long  (read-only)`
- `Text As String  (read/write)`
- `Count As Long  (read-only)`
- `_NewEnum As IUnknown  (read-only)`
- `Parent As Object  (read-only)`
- `Paragraphs As TextRange2  (read-only)`
- `Sentences As TextRange2  (read-only)`
- `Words As TextRange2  (read-only)`
- `Characters As TextRange2  (read-only)`
- `Lines As TextRange2  (read-only)`
- `Runs As TextRange2  (read-only)`
- `ParagraphFormat As ParagraphFormat2  (read-only)`
- `Font As Font2  (read-only)`
- `Length As Long  (read-only)`
- `Start As Long  (read-only)`
- `BoundLeft As Single  (read-only)`
- `BoundTop As Single  (read-only)`
- `BoundWidth As Single  (read-only)`
- `BoundHeight As Single  (read-only)`
- `LanguageID As MsoLanguageID  (read/write)`
- `MathZones As TextRange2  (read-only)`

## Methods (20)

- `Item(Index As Variant) As TextRange2`
- `TrimText() As TextRange2`
- `InsertAfter([NewText As String]) As TextRange2`
- `InsertBefore([NewText As String]) As TextRange2`
- `InsertSymbol(FontName As String, CharNumber As Long, [Unicode As MsoTriState]) As TextRange2`
- `Select()`
- `Cut()`
- `Copy()`
- `Delete()`
- `Paste() As TextRange2`
- `PasteSpecial(Format As MsoClipboardFormat) As TextRange2`
- `ChangeCase(Type As MsoTextChangeCase)`
- `AddPeriods()`
- `RemovePeriods()`
- `Find(FindWhat As String, [After As Long], [MatchCase As MsoTriState], [WholeWords As MsoTriState]) As TextRange2`
- `Replace(FindWhat As String, ReplaceWhat As String, [After As Long], [MatchCase As MsoTriState], [WholeWords As MsoTriState]) As TextRange2`
- `RotatedBounds(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, X3 As Single, Y3 As Single, x4 As Single, y4 As Single)`
- `RtlRun()`
- `LtrRun()`
- `InsertChartField(ChartFieldType As MsoChartFieldType, [Formula As String], [Position As Long]) As TextRange2`
