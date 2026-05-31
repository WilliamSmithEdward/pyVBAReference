# SharedWorkspaceLink

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C037F-0000-0000-C000-000000000046}  

Represents a URL link saved in a shared document workspace site.

**Remarks:** Use the SharedWorkspaceLink object to manage links to additional documents and information of interest to the members who are collaborating on the documents in the shared workspace site. Use the Item (_index_) property of the SharedWorkspaceLinks collection to return a specific SharedWorkspaceLink object. Use the Description property to set the link description that appears on the Links tab of the Shared Workspace pane and on the workspace webpage. Use the URL property to set the destination address of the link. Use the Notes property to supply additional information about the link. Use the Save method to upload changes to the server after you modify properties of the SharedWorkspaceLink object. Use the CreatedBy, CreatedDate, ModifiedBy, and ModifiedDate properties to return information about the history of each link.

**Example:**

```vba
Dim swsLink As Office.SharedWorkspaceLink
    Set swsLink = ActiveWorkbook.SharedWorkspace.Links(1)
    With swsLink
        .Description = "MSDN Home Page"
        .URL = "https://msdn.microsoft.com/"
        .Notes = "My favorite site for developers!"
        .Save
    End With
    Set swsLink = Nothing
```

## Properties (10)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the SharedWorkspaceFile object (you can use this property with an Automation object to return that object's container application). Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the SharedWorkspaceLink object was created. Read-only.
- `URL As String  (read/write)`  
  Gets the top-level Uniform Resource Locator (URL) of the shared workspace link. Read/write.
- `Description As String  (read/write)`  
  Gets or sets a descriptive String value for the specified SharedWorkspaceLink or SharedWorkspaceTask object. Read/write.
- `Notes As String  (read/write)`  
  Gets or sets the optional notes associated with a shared workspace link. Read/write.
- `CreatedBy As String  (read-only)`  
  Gets the display name of the member who created the shared workspace object. Read-only.
- `CreatedDate As Variant  (read-only)`  
  Gets the date and time when the shared workspace object was created. Read-only.
- `ModifiedBy As String  (read-only)`  
  Gets the name of the user who last modified the object. Read-only.
- `ModifiedDate As Variant  (read-only)`  
  Gets the date and time when the SharedWorkspaceLink object was last modified. Read-only.
- `Parent As Object  (read-only)`  
  Gets the Parent object for the SharedWorkspaceLink object. Read-only.

## Methods (2)

- `Save()`  
  Uploads changes made programmatically to a shared server.
- `Delete()`  
  Deletes the current SharedWorkspaceLink object.
