# DocumentLibraryVersions

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C0388-0000-0000-C000-000000000046}  

The DocumentLibraryVersions property of the Document object in Microsoft Word, the Workbook object in Excel, and the Presentation object in PowerPoint returns a DocumentLibraryVersions object. The DocumentLibraryVersions object represents a collection of DocumentLibraryVersion objects.

**Remarks:** Use the DocumentLibraryVersions object with documents stored in a SharePoint document library on the server to determine whether versioning is enabled for the active document, and if versioning is enabled, to manage the document's collection of DocumentLibraryVersion objects. Each DocumentLibraryVersion object represents one saved version of the active document. When versioning is enabled, a new version is created on the server when the following actions occur; additional versions are not created each time the user saves changes to the open document. - Check in - Save: A new version is created on the server when the user first saves the document after opening it. Additional changes saved while the document is open apply to the same version. - Restore - Upload The DocumentLibraryVersions object model is available whether versioning is enabled or disabled on the active document. The DocumentLibraryVersions property of the Document, Workbook, and Presentation objects does not return Nothing when the active document is not stored in a document library or versioning is not enabled.

**Example:**

```vba
Dim dlvVersions As Office.DocumentLibraryVersions
 Dim dlvVersion As Office.DocumentLibraryVersion
 Dim strVersionInfo As String
 Set dlvVersions = ActiveDocument.DocumentLibraryVersions
 If dlvVersions.IsVersioningEnabled Then
 strVersionInfo = "This document has " & _
 dlvVersions.Count & " versions: " & vbCrLf
 For Each dlvVersion In dlvVersions
 strVersionInfo = strVersionInfo & _
 " - Version #: " & dlvVersion.Index & vbCrLf & _
 " - Modified by: " & dlvVersion.ModifiedBy & vbCrLf & _
 " - Modified on: " & dlvVersion.Modified & vbCrLf & _
 " - Comments: " & dlvVersion.Comments & vbCrLf
 Next
 Else
 strVersionInfo = "Versioning not enabled for this document."
 End If
 MsgBox strVersionInfo, vbInformation + vbOKOnly, "Version Information"
 Set dlvVersion = Nothing
 Set dlvVersions = Nothing
```

## Properties (7)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the DocumentLibraryVersions object (you can use this property with an Automation object to return that object's container application). Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the DocumentLibraryVersions object was created. Read-only.
- `Item As DocumentLibraryVersion  (read-only)`  
  Gets a DocumentLibraryVersion object from the DocumentLibraryVersions collection. Read-only.
- `Count As Long  (read-only)`  
  Gets a Long indicating the number of items in the DocumentLibraryVersions collection. Read-only.
- `Parent As Object  (read-only)`  
  Gets the Parent object for the DocumentLibraryVersions object. Read-only.
- `IsVersioningEnabled As Boolean  (read-only)`  
  Gets a Boolean value that indicates whether the document library in which the active document is saved on the server is configured to create a backup copy, or version, each time the file is edited on the website. Read-only.
- `_NewEnum As IUnknown  (read-only)`
