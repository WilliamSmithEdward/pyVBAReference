# Bookmark

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020968-0000-0000-C000-000000000046}  

Represents a single bookmark in a document, selection, or range. The Bookmark object is a member of the Bookmarks collection. The Bookmarks collection includes all the bookmarks listed in the Bookmark dialog box (Insert menu).

**Remarks:** Using the Bookmark Object Use Bookmarks (_index_), where _index_ is the bookmark name or index number, to return a single Bookmark object. You must exactly match the spelling (but not necessarily the capitalization) of the bookmark name. The following example selects the bookmark named "temp" in the active document. The index number represents the position of the bookmark in the Selection or Range object. For the Document object, the index number represents the position of the bookmark in the alphabetical list of bookmarks in the Bookmarks dialog box (click Name to sort the list of bookmarks alphabetically). The following example displays the name of the second bookmark in the Bookmarks collection. Use the Add method to add a bookmark to a document range. The following example marks the selection by adding a bookmark named "temp." Remarks Use the BookmarkID property with a range or selection object to return the index number of a Bookmark object in the Bookmarks collection. The following example displays the index number of the bookmark named "temp" in the active document. Use predefined bookmarkswith the Bookmarks property.

## Properties (10)

- `Name As String  (read-only)`  
  Returns the name of the specified object. Read-only String.
- `Range As Range  (read-only)`  
  Returns a Range object that represents the portion of a document that's contained in the specified object.
- `Empty As Boolean  (read-only)`  
  True if the specified bookmark is empty. Read-only Boolean.
- `Start As Long  (read/write)`  
  Returns or sets the starting character position of a bookmark. Read/write Long.
- `End As Long  (read/write)`  
  Returns or sets the ending character position of a selection, range, or bookmark. Read/write Long.
- `Column As Boolean  (read-only)`  
  True if the specified bookmark is a table column. Read-only Boolean.
- `StoryType As WdStoryType  (read-only)`  
  Returns the story type for the specified range, selection, or bookmark. Read-only WdStoryType.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Bookmark object.

## Methods (3)

- `Select()`  
  Selects the specified bookmark.
- `Delete()`  
  Deletes the specified bookmark.
- `Copy(Name As String) As Bookmark`  
  Copies a bookmark to the new bookmark specified in the Name argument, and returns a Bookmark object.
    - `Name As String` (required): The name of the new bookmark.
