# CommentsThreaded

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244FC-0000-0000-C000-000000000046}  

A collection of top-level CommentThreaded objects in a Worksheet, or a collection of replies in a single threaded comment.

**Remarks:** Each threaded comment is represented by a CommentThreaded object.

**Example:**

```vba
Set cmt = Worksheets(1).CommentsThreaded
For Each c In cmt
 c.Text "Updated Comment"
Next
```

## Properties (6)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Count As Long  (read-only)`  
  Returns a Long value that represents the number of objects in the collection.
- `_Default As CommentThreaded  (read-only)`
- `_NewEnum As IUnknown  (read-only)`

## Methods (1)

- `Item(Index As Long) As CommentThreaded`  
  Returns a single object from a collection.
    - `Index As Long` (required): The index number for the object.
