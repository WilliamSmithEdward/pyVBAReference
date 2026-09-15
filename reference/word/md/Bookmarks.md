# Bookmarks

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020967-0000-0000-C000-000000000046}  

A collection of Bookmark objects that represent the bookmarks in the specified selection, range, or document.

**Remarks:** Use the Bookmarks property to return the Bookmarks collection for a document, range, or selection. The following example ensures that the bookmark named "temp" exists in the active document before selecting the bookmark. Use the Add method to set a bookmark for a range in a document. The following example marks the selection by adding a bookmark named "temp". Use Bookmarks (index), where index is the bookmark name or index number, to return a single Bookmark object. You must exactly match the spelling (but not necessarily the capitalization) of the bookmark name. The following example selects the bookmark named "temp" in the active document. The index number represents the position of the bookmark in the Selection or Range object. For the Document object, the index number represents the position of the bookmark in the alphabetical list of bookmarks in the Bookmarks dialog box (click Name to sort the list of bookmarks alphabetically). The following example displays the name of the second bookmark in the Bookmarks collection. Remarks The ShowHidden property effects the number of elements in the Bookmarks collection.

## Properties (7)

- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns the number of items in the Bookmarks collection. Read-only Long.
- `DefaultSorting As WdBookmarkSortBy  (read/write)`  
  Returns or sets the sorting option for bookmark names displayed in the Bookmark dialog box (Insert menu). Read/write WdBookmarkSortBy.
- `ShowHidden As Boolean  (read/write)`  
  True if hidden bookmarks are included in the Bookmarks collection. Read/write Boolean.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Bookmarks collection.

## Methods (3)

- `Item(Index As Variant) As Bookmark`  
  Returns an individual Bookmark object in a collection.
    - `Index As Variant` (required): The individual object to be returned. Can be a Long indicating the ordinal position or a String representing the name of the individual object.
- `Add(Name As String, [Range As Variant]) As Bookmark`  
  Returns a Bookmark object that represents a bookmark added to a range.
    - `Name As String` (required): The name of the bookmark. The name cannot be more than 40 characters or include more than one word.
    - `Range As Variant` (optional): The range of text marked by the bookmark. A bookmark can be set to a collapsed range (the insertion point).
- `Exists(Name As String) As Boolean`  
  Determines whether the specified bookmark exists. Returns True if the bookmark exists.
    - `Name As String` (required): A bookmark name than can not include more than 40 characters or more than one word.
