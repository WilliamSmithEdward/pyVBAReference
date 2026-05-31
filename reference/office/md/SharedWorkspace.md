# SharedWorkspace

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C0385-0000-0000-C000-000000000046}  

The SharedWorkspace property of a Document object in Microsoft Word, a Workbook object in Microsoft Excel, and a Presentation object in Microsoft PowerPoint returns a SharedWorkspace object that allows the developer to add the active document to a SharePoint site and to manage other objects in the shared workspace site.

**Remarks:** Use the SharedWorkspace object to add the active Word, Excel, or PowerPoint document to a SharePoint document workspace site on the server to take advantage of the workspace's collaboration features, or to disconnect or remove the document from the workspace site. Use the SharedWorkspace object's collections to manage files, folders, links, members, and tasks associated with the shared document. The SharedWorkspace object model is available whether or not a document is stored in a workspace. The SharedWorkspace property of the Document, Workbook, and Presentation objects does not return Nothing when the document is not shared. Use the Connected property of the SharedWorkspace object to determine whether the active document is in fact saved in and connected to a shared workspace. Users require appropriate permissions to use the objects, properties, and methods in the SharedWorkspace object hierarchy. Use the SharedWorkspaceFiles collection, accessed through the Files property of the SharedWorkspace object, to manage documents and files saved in a shared workspace.

**Example:**

```vba
Dim swsWorkspace As Office.SharedWorkspace
    Dim strSWSInfo As String
    Set swsWorkspace = ActiveWorkbook.SharedWorkspace
    strSWSInfo = swsWorkspace.Name & vbCrLf & _
        " - URL: " & swsWorkspace.URL & vbCrLf & _
        "The shared workspace contains " & vbCrLf & _
        " - Files: " & swsWorkspace.Files.Count & vbCrLf & _
        " - Folders: " & swsWorkspace.Folders.Count & vbCrLf & _
        " - Links: " & swsWorkspace.Links.Count & vbCrLf & _
        " - Members: " & swsWorkspace.Members.Count & vbCrLf & _
        " - Tasks: " & swsWorkspace.Tasks.Count & vbCrLf
    MsgBox strSWSInfo, vbInformation + vbOKOnly, _
        "Shared Workspace Information"
    Set swsWorkspace = Nothing
```

## Properties (13)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the SharedWorkspace object (you can use this property with an Automation object to return that object's container application). Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the SharedWorkspace object was created. Read-only.
- `Name As String  (read/write)`  
  Gets or sets the display name of the shared workspace site. Read/write.
- `Members As SharedWorkspaceMembers  (read-only)`  
  Gets a SharedWorkspaceMembers collection that represents the list of members in the current shared workspace. Read-only.
- `Tasks As SharedWorkspaceTasks  (read-only)`  
  Gets a SharedWorkspaceTasks collection that represents the list of tasks in the current shared workspace. Read-only.
- `Files As SharedWorkspaceFiles  (read-only)`  
  Provides access to the SharedWorkspaceFile objects in the SharedWorkspace. Read-only.
- `Folders As SharedWorkspaceFolders  (read-only)`  
  Gets a SharedWorkspaceFolders collection that represents the list of subfolders in the document library associated with the current shared workspace. Read-only.
- `Links As SharedWorkspaceLinks  (read-only)`  
  Gets a SharedWorkspaceLinks collection that represents the list of links saved in the current shared workspace. Read-only.
- `Parent As Object  (read-only)`  
  Gets the Parent object for the SharedWorkspace object. Read-only.
- `URL As String  (read-only)`  
  Gets the top-level Uniform Resource Locator (URL) of the shared workspace. Read-only.
- `Connected As Boolean  (read-only)`  
  Gets a Boolean value that indicates whether or not the active document is currently saved in and connected to a shared workspace. Read-only.
- `LastRefreshed As Variant  (read-only)`  
  Gets the date and time when the Refresh method was most recently called. Read-only.
- `SourceURL As String  (read/write)`  
  Designates the location of the public copy of a shared document to which changes should be published back after the document has been revised in a separate document workspace site. Read-only.

## Methods (5)

- `Refresh()`  
  Refreshes the local cache of the SharedWorkspace object's files, folders, links, members, and tasks from the server.
- `CreateNew([URL As Variant], [Name As Variant])`  
  Creates a document workspace site on the server and adds the active document to the new shared workspace site.
    - `URL As Variant` (optional): The URL for the parent folder in which the new shared workspace is to be created. If you don't supply a URL, the site is created in the user's default server location.
    - `Name As Variant` (optional): The name of the new shared workspace site. The default value is the name of the active document without its file name extension. For example, if you create a workspace site for "Budget.xls", the name of the new site becomes "Budget".
- `Delete()`  
  Deletes the current shared workspace and all data within it.
- `RemoveDocument()`  
  Removes the active document from the shared workspace site.
- `Disconnect()`  
  Disconnects the local copy of the active document from the shared workspace site.
