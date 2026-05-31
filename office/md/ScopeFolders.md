# ScopeFolders

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C0369-0000-0000-C000-000000000046}  

A collection of ScopeFolder objects.

**Remarks:** Only ScopeFolder objects contain ScopeFolders collections. Each ScopeFolders collection contains the ScopeFolder objects that correspond to the subfolders of the parent ScopeFolder object. Use the ScopeFolders property of the ScopeFolder object to return a ScopeFolders collection. You can't add or remove ScopeFolder objects from a ScopeFolders collection.

## Properties (5)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the ScopeFolders object (you can use this property with an Automation object to return that object's container application). Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the ScopeFolders object was created. Read-only.
- `Item As ScopeFolder  (read-only)`  
  Gets a ScopeFolder object that represents a subfolder of the parent object. Read-only.
- `Count As Long  (read-only)`  
  Gets a Long indicating the number of items in the ScopeFolders collection. Read-only.
- `_NewEnum As IUnknown  (read-only)`
