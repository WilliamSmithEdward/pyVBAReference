# DocumentEvents2

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020A02-0000-0000-C000-000000000046}  

## Methods (14)

- `New()`
- `Open()`
- `Close()`
- `Sync(SyncEventType As MsoSyncEventType)`
- `XMLAfterInsert(NewXMLNode As XMLNode, InUndoRedo As Boolean)`
- `XMLBeforeDelete(DeletedRange As Range, OldXMLNode As XMLNode, InUndoRedo As Boolean)`
- `ContentControlAfterAdd(NewContentControl As ContentControl, InUndoRedo As Boolean)`
- `ContentControlBeforeDelete(OldContentControl As ContentControl, InUndoRedo As Boolean)`
- `ContentControlOnExit(ContentControl As ContentControl, Cancel As Boolean)`
- `ContentControlOnEnter(ContentControl As ContentControl)`
- `ContentControlBeforeStoreUpdate(ContentControl As ContentControl, Content As String)`
- `ContentControlBeforeContentUpdate(ContentControl As ContentControl, Content As String)`
- `BuildingBlockInsert(Range As Range, Name As String, Category As String, BlockType As String, Template As String)`
- `ContentControlNonContentChange(ContentControl As ContentControl)`
