# RecentFiles

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020963-0000-0000-C000-000000000046}  

A collection of RecentFile objects that represents the files that have been used recently. The items in the RecentFiles collection are displayed at the bottom of the File menu.

**Remarks:** Use the RecentFiles property to return the RecentFiles collection. The following example sets five as the maximum number of files that the RecentFiles collection can contain. Use the Add method to add a file to the RecentFiles collection. The following example adds the active document to the list of recently-used files. Use RecentFiles (Index), where Index is the index number, to return a single RecentFile object. The index number represents the position of the file on the File menu. The following example opens the first document in the RecentFiles collection. The SaveAs and Open methods include an AddToRecentFiles argument that controls whether or not a file is added to the recently-used-files list when the file is opened or saved.

## Properties (6)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified RecentFiles object.
- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of recent files in the collection. Read-only.
- `Maximum As Long  (read/write)`  
  Returns or sets the maximum number of recently used files that can appear on the File menu. Can be a number from 0 (zero) through 9. Read/write Long.

## Methods (2)

- `Item(Index As Long) As RecentFile`  
  Returns an individual RecentFile object in a collection.
    - `Index As Long` (required): The individual object to be returned. Can be a Long indicating the ordinal position of the individual object.
- `Add(Document As Variant, [ReadOnly As Variant]) As RecentFile`  
  Returns a RecentFile object that represents a file added to the list of recently used files.
    - `Document As Variant` (required): The document you want to add to the list of recently used files. You can specify this argument by using either the string name for the document or a Document object.
    - `ReadOnly As Variant` (optional): True to make the document read-only.
