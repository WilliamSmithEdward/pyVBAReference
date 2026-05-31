# ScopeFolder

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C0368-0000-0000-C000-000000000046}  

Corresponds to a searchable folder. ScopeFolder objects are intended for use with the SearchFolders collection.

**Remarks:** When you want to search specific folders, you can use the methods and properties of the SearchScope object and ScopeFolders collection to retrieve ScopeFolder objects and add them to the SearchFolders collection. In each ScopeFolder object, there is a ScopeFolders collection that contains the subfolders of the parent ScopeFolder object. You can traverse the entire folder structure of a search scope (for example, all local drives) by looping through these ScopeFolders collections and returning all of the lower-level ScopeFolder objects. A ScopeFolder object with no subfolders contains an empty ScopeFolders collection. For an example that demonstrates how to loop through all of the ScopeFolder objects in a search scope, see the SearchFolders collection topic. Use the Add method of the SearchFolders collection to add a ScopeFolder object to the SearchFolders collection; however, it is usually simpler to use the AddToSearchFolders method of the ScopeFolder that you want to add because there is only one SearchFolders collection for all searches. For an example that demonstrates how to add a ScopeFolder to the SearchFolders collection, see the SearchFolders collection topic.

**Example:**

```vba
Set sf = SearchScopes.Item(1).ScopeFolder
```

## Properties (5)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the ScopeFolder object (you can use this property with an Automation object to return that object's container application). Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the ScopeFolder object was created. Read-only.
- `Name As String  (read-only)`  
  Gets the name of a searchable folder. Read-only.
- `Path As String  (read-only)`  
  Gets a String indicating the full path of a ScopeFolder object. Read-only.
- `ScopeFolders As ScopeFolders  (read-only)`  
  Gets a ScopeFolders collection. The items in this collection correspond to the subfolders of the parent ScopeFolder object. Read-only.

## Methods (1)

- `AddToSearchFolders()`  
  Adds a ScopeFolder object to the SearchFolders collection.
