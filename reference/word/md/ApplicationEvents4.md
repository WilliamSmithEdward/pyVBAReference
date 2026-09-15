# ApplicationEvents4

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020A01-0000-0000-C000-000000000046}  

## Methods (34)

- `Quit()`
- `DocumentChange()`
- `DocumentOpen(Doc As Document)`
- `DocumentBeforeClose(Doc As Document, Cancel As Boolean)`
- `DocumentBeforePrint(Doc As Document, Cancel As Boolean)`
- `DocumentBeforeSave(Doc As Document, SaveAsUI As Boolean, Cancel As Boolean)`
- `NewDocument(Doc As Document)`
- `WindowActivate(Doc As Document, Wn As Window)`
- `WindowDeactivate(Doc As Document, Wn As Window)`
- `WindowSelectionChange(Sel As Selection)`
- `WindowBeforeRightClick(Sel As Selection, Cancel As Boolean)`
- `WindowBeforeDoubleClick(Sel As Selection, Cancel As Boolean)`
- `EPostagePropertyDialog(Doc As Document)`
- `EPostageInsert(Doc As Document)`
- `MailMergeAfterMerge(Doc As Document, DocResult As Document)`
- `MailMergeAfterRecordMerge(Doc As Document)`
- `MailMergeBeforeMerge(Doc As Document, StartRecord As Long, EndRecord As Long, Cancel As Boolean)`
- `MailMergeBeforeRecordMerge(Doc As Document, Cancel As Boolean)`
- `MailMergeDataSourceLoad(Doc As Document)`
- `MailMergeDataSourceValidate(Doc As Document, Handled As Boolean)`
- `MailMergeWizardSendToCustom(Doc As Document)`
- `MailMergeWizardStateChange(Doc As Document, FromState As Long, ToState As Long, Handled As Boolean)`
- `WindowSize(Doc As Document, Wn As Window)`
- `XMLSelectionChange(Sel As Selection, OldXMLNode As XMLNode, NewXMLNode As XMLNode, Reason As Long)`
- `XMLValidationError(XMLNode As XMLNode)`
- `DocumentSync(Doc As Document, SyncEventType As MsoSyncEventType)`
- `EPostageInsertEx(Doc As Document, cpDeliveryAddrStart As Long, cpDeliveryAddrEnd As Long, cpReturnAddrStart As Long, cpReturnAddrEnd As Long, xaWidth As Long, yaHeight As Long, bstrPrinterName As String, bstrPaperFeed As String, fPrint As Boolean, fCancel As Boolean)`
- `MailMergeDataSourceValidate2(Doc As Document, Handled As Boolean)`
- `ProtectedViewWindowOpen(PvWindow As ProtectedViewWindow)`
- `ProtectedViewWindowBeforeEdit(PvWindow As ProtectedViewWindow, Cancel As Boolean)`
- `ProtectedViewWindowBeforeClose(PvWindow As ProtectedViewWindow, CloseReason As Long, Cancel As Boolean)`
- `ProtectedViewWindowSize(PvWindow As ProtectedViewWindow)`
- `ProtectedViewWindowActivate(PvWindow As ProtectedViewWindow)`
- `ProtectedViewWindowDeactivate(PvWindow As ProtectedViewWindow)`
