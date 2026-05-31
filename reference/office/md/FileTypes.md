# FileTypes

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C036C-0000-0000-C000-000000000046}  

A collection of values of the type msoFileType that determine which types of files are returned during a search.

**Remarks:** There is only one FileTypes collection for all searches, so it's important to clear the FileTypes collection before executing a search unless you wish to search for file types from previous searches. The easiest way to clear the collection is to set the FileType property to the first file type for which you want to search. You can also remove individual types by using the Remove method. To determine the file type of each item in the collection, use the Item property to return the msoFileType value.

## Properties (5)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the FileTypes object (you can use this property with an Automation object to return that object's container application). Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the FileTypes object was created. Read-only.
- `Item As MsoFileType  (read-only)`  
  Gets a value that indicates which file type will be searched for. Read-only.
- `Count As Long  (read-only)`  
  Gets a Long indicating the number of items in the FileTypes collection. Read-only.
- `_NewEnum As IUnknown  (read-only)`

## Methods (2)

- `Add(FileType As MsoFileType)`  
  Adds a new file type to a file search.
    - `FileType As MsoFileType` (required): Specifies the type of file for which to search.
- `Remove(Index As Long)`  
  Removes a FileType object from the collection.
    - `Index As Long` (required): The index number of the file type to be removed.
