# SearchFolders

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C036A-0000-0000-C000-000000000046}  

A collection of ScopeFolder objects that determines which folders are searched.

**Remarks:** For each application, there is only a single SearchFolders collection. The contents of the collection remains after the code that calls it has finished executing. Consequently, it is important to clear the collection unless you want to include folders from previous searches in your search. Use the Add method of the SearchFolders collection to add a ScopeFolder object to the SearchFolders collection; however, it is usually simpler to use the AddToSearchFolders method of the ScopeFolder that you want to add because there is only one SearchFolders collection for all searches.

## Properties (5)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the SearchFolders object (you can use this property with an Automation object to return that object's container application). Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the SearchFolders object was created. Read-only.
- `Item As ScopeFolder  (read-only)`  
  Gets a ScopeFolder object that represents a subfolder of the parent object. Read-only.
- `Count As Long  (read-only)`  
  Gets a Long indicating the number of items in the SearchFolders collection. Read-only.
- `_NewEnum As IUnknown  (read-only)`

## Methods (2)

- `Add(ScopeFolder As ScopeFolder)`  
  Adds a search folder to a file search.
    - `ScopeFolder As ScopeFolder` (required): The folder to add to the search.
- `Remove(Index As Long)`  
  Removes the specified object from the collection.
    - `Index As Long` (required): The index number of the folder to be removed.
