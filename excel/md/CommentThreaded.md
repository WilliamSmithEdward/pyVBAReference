# CommentThreaded

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244FD-0000-0000-C000-000000000046}  

Represents a cell's threaded comment. This object can represent both a top-level comment or its replies.

**Remarks:** The CommentThreaded object is a member of the CommentsThreaded collection.

**Example:**

```vba
Worksheets(1).Range("E5").CommentThreaded.Text "reviewed on " & Date
```

## Properties (7)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Replies As CommentsThreaded  (read-only)`  
  If this comment is a parent, returns a CommentsThreaded collection of CommentThreaded objects that are children/replies of the specified comment (if any exist). The replies are sorted by time stamp.
- `Author As Author  (read-only)`  
  Returns the Author object that represents the author of the specified CommentThreaded object. Read-only.
- `Date As Variant  (read-only)`  
  Returns a date String-Variant that represents the date and time that a threaded comment was added in local time. Read-only.
- `Resolved As Boolean  (read/write)`

## Methods (5)

- `AddReply([Text As Variant]) As CommentThreaded`  
  If the comment is a top-level comment, it will add a reply to its replies collection.
    - `Text As Variant` (optional): The reply's text.
- `Delete()`  
  Deletes the specified threaded comment and all replies associated with that comment (if any exist).
- `Text([Text As Variant], [Start As Variant], [Overwrite As Variant]) As String`  
  Sets threaded comment text.
    - `Text As Variant` (optional): The text to be added.
    - `Start As Variant` (optional): The character number where the added text will be placed. If the _Overwrite_ parameter is True or blank, and if this argument is omitted, any existing text in the threaded comment is deleted.
    - `Overwrite As Variant` (optional): False to insert the text. The default value is True (text is overwritten).
- `Next() As CommentThreaded`  
  Returns a CommentThreaded object that represents the next threaded comment.
- `Previous() As CommentThreaded`  
  Returns a CommentThreaded object that represents the previous threaded comment.
