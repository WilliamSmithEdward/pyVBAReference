# MediaBookmarks

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {BA72E556-4FF5-48F4-8215-5505F990966F}  

A collection of MediaBookmark objects.

## Properties (1)

- `Count As Long  (read-only)`  
  Returns the number of bookmarks. Read-only.

## Methods (2)

- `Item(Index As Long) As MediaBookmark`  
  Returns the specified MediaBookmark.
    - `Index As Long` (required): The position of the MediaBookmark.
- `Add(Position As Long, Name As String) As MediaBookmark`  
  Adds a new MediaBookmark at the specified time and using the specified name.
    - `Position As Long` (required): The position of the MediaBookmark.
    - `Name As String` (required): The name of the MediaBookmark.
