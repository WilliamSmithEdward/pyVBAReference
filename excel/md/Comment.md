# Comment

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024427-0000-0000-C000-000000000046}  

Represents a cell comment.

**Remarks:** The Comment object is a member of the Comments collection.

**Example:**

```vba
Worksheets(1).Range("E5").Comment.Text "reviewed on " & Date
```

## Properties (6)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Author As String  (read-only)`  
  Returns the author of the comment. Read-only String.
- `Shape As Shape  (read-only)`  
  Returns a Shape object that represents the shape attached to the specified comment.
- `Visible As Boolean  (read/write)`  
  Returns or sets a Boolean value that determines whether the object is visible. Read/write.

## Methods (4)

- `Text([Text As Variant], [Start As Variant], [Overwrite As Variant]) As String`  
  Sets comment text.
    - `Text As Variant` (optional): The text to be added.
    - `Start As Variant` (optional): The character number where the added text will be placed. If this argument is omitted, any existing text in the comment is deleted.
    - `Overwrite As Variant` (optional): False to insert the text. The default value is True (text is overwritten).
- `Delete()`  
  Deletes the object.
- `Next() As Comment`  
  Returns a Comment object that represents the next comment.
- `Previous() As Comment`  
  Returns a Comment object that represents the previous comment.
