# CommentThreaded

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244FD-0000-0000-C000-000000000046}  

## Properties (7)

- `Application As Application  (read-only)`
- `Creator As XlCreator  (read-only)`
- `Parent As Object  (read-only)`
- `Replies As CommentsThreaded  (read-only)`
- `Author As Author  (read-only)`
- `Date As Variant  (read-only)`
- `Resolved As Boolean  (read/write)`

## Methods (5)

- `AddReply([Text As Variant]) As CommentThreaded`
- `Delete()`
- `Text([Text As Variant], [Start As Variant], [Overwrite As Variant]) As String`
- `Next() As CommentThreaded`
- `Previous() As CommentThreaded`
