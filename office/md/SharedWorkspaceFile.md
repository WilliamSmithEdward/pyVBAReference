# SharedWorkspaceFile

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C037B-0000-0000-C000-000000000046}  

The SharedWorkspaceFile object represents a file saved in a shared document workspace.

**Remarks:** Use the SharedWorkspaceFile object to manage documents and files saved in a shared workspace.

**Example:**

```vba
Private Function FilenameFromURL(FileURL As String) As String
    Dim intLastSeparator As Integer
    FileURL = URLDecode(FileURL)
    intLastSeparator = InStrRev(FileURL, "/")
    FilenameFromURL = Right(FileURL, Len(FileURL) - intLastSeparator)
End Function

Private Function URLDecode(URLtoDecode As String) As String
    URLDecode = Replace(URLtoDecode, "%20", " ")
End Function
```

## Properties (8)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the SharedWorkspaceFile object (you can use this property with an Automation object to return that object's container application). Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the SharedWorkspaceFile object was created. Read-only.
- `URL As String  (read-only)`  
  Gets the full Uniform Resource Locator (URL) and file name of the shared workspace file. Read-only.
- `CreatedBy As String  (read-only)`  
  Gets the display name of the member who created the shared workspace object. Read-only.
- `CreatedDate As Variant  (read-only)`  
  Gets the date and time when the shared workspace object was created. Read-only.
- `ModifiedBy As String  (read-only)`  
  Gets the name of the user who last modified the object. Read-only.
- `ModifiedDate As Variant  (read-only)`  
  Gets the date and time when the SharedWorkspaceFile object was last modified. Read-only.
- `Parent As Object  (read-only)`  
  Gets the Parent object for the SharedWorkspaceFile object. Read-only.

## Methods (1)

- `Delete()`  
  Deletes the current SharedWorkspaceFile object.
