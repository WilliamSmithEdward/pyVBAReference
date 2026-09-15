# IApplicationEvents3

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020A00-0001-0000-C000-000000000046}  

## Methods (23)

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
