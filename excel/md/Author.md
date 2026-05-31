# Author

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244FE-0000-0000-C000-000000000046}  

Represents the author of the CommentThreaded object.

**Remarks:** Use the Author property of the CommentThreaded object to return the Author object.

**Example:**

```vba
Worksheets(1).Range("A1").CommentThreaded.Author.Name
```

## Properties (6)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Name As String  (read-only)`  
  Returns a String that contains the display name of the specified comment author. Read-only.
- `ProviderID As String  (read-only)`  
  Returns a String that represents the ID of the service providing the contact information. Read-only.
- `UserID As String  (read-only)`  
  Returns a String that represents the user ID of the contact. Read-only.
