# RecentFile

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020964-0000-0000-C000-000000000046}  

Represents a recently used file. The RecentFile object is a member of the RecentFiles collection.

**Remarks:** The RecentFiles collection includes all the files that have been used recently. The items in the RecentFiles collection are displayed at the bottom of the File menu. Use RecentFiles (Index), where Index is the index number, to return a single RecentFile object. The index number represents the position of the file on the File menu. The following example opens the first document in the RecentFiles collection. Use the Add method to add a file to the RecentFiles collection. The following example adds the active document to the list of recently-used files. The SaveAs and Open methods include an AddToRecentFiles argument that controls whether or not a file is added to the recently-used-files list when the file is opened or saved.

## Properties (7)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified RecentFile object.
- `Name As String  (read-only)`  
  Returns the name of the specified object. Read-only String.
- `Index As Long  (read-only)`  
  Returns a Long that represents the position of an item in a collection. Read-only.
- `ReadOnly As Boolean  (read/write)`  
  True if changes to the document cannot be saved to the original document. Read/write Boolean.
- `Path As String  (read-only)`  
  Returns the disk or Web path to the specified object. Read-only String.

## Methods (2)

- `Open() As Document`  
  Opens the specified object. Returns a Document object representing the opened document.
- `Delete()`  
  Deletes the specified file on the list of recent files.
