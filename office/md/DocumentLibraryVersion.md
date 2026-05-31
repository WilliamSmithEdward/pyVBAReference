# DocumentLibraryVersion

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C0387-0000-0000-C000-000000000046}  

The DocumentLibraryVersion object represents a single saved version of a shared document that has versioning enabled and which is stored in a document library on the server. Each DocumentLibraryVersion object is a member of the active document's DocumentLibraryVersions collection.

**Remarks:** Each DocumentLibraryVersion object represents one saved version of the active document. When versioning is enabled, a new version is created on the server when the actions listed below occur; additional versions are not created each time the user saves changes to the open document. - Check in - Save: A new version is created on the server when the user first saves the document after opening it. Additional changes saved while the document is open apply to the same version. - Restore - Upload Use the Modified, ModifiedBy, and Comments properties to return information about a saved version of a shared document. Use the Open method to open a previous version, or the Restore method to restore a previous version in place of the current version. Use the Delete method to delete a version.

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
  Gets an Application object that represents the container application for the DocumentLibraryVersion object (you can use this property with an Automation object to return that object's container application). Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the DocumentLibraryVersion object was created. Read-only.
- `Modified As Variant  (read-only)`  
  Gets the date and time at which the specified version of the shared document was last saved to the server. Read-only.
- `Index As Long  (read-only)`  
  Gets a Long representing the index number for a DocumentLibraryVersion object in the collection. Read-only.
- `Parent As Object  (read-only)`  
  Gets the Parent object for the DocumentLibraryVersion object. Read-only.
- `ModifiedBy As String  (read-only)`  
  Gets the name of the user who last saved the specified version of the shared document to the server. Read-only.
- `Comments As String  (read-only)`  
  Gets any optional comments associated with the specified version of the shared document. Read-only.

## Methods (3)

- `Delete()`  
  Removes a document library version from the DocumentLibraryVersions collection.
- `Open() As Object`  
  Opens the specified version of the shared document from the DocumentLibraryVersions collection in read-only mode.
- `Restore() As Object`  
  Restores a previous saved version of a shared document from the DocumentLibraryVersions collection.
