# ICommentThreaded

**Type:** Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244FD-0001-0000-C000-000000000046}  

## Properties (7)

- `Application As HRESULT  (read-only)`
- `Creator As HRESULT  (read-only)`
- `Parent As HRESULT  (read-only)`
- `Replies As HRESULT  (read-only)`
- `Author As HRESULT  (read-only)`
- `Date As HRESULT  (read-only)`
- `Resolved As HRESULT  (read/write)`

## Methods (5)

- `AddReply([Text As Variant], RHS As CommentThreaded)`
- `Delete()`
- `Text([Text As Variant], [Start As Variant], [Overwrite As Variant], RHS As String)`
- `Next(RHS As CommentThreaded)`
- `Previous(RHS As CommentThreaded)`
