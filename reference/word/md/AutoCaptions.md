# AutoCaptions

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {0002097A-0000-0000-C000-000000000046}  

A collection of AutoCaption objects that represent the captions that can be automatically added when items such as tables, pictures, or OLE objects are inserted into a document.

**Remarks:** Use the AutoCaptions property to return the AutoCaptions collection. The following example displays the names of the selected items in the AutoCaption dialog box. The AutoCaptions collection contains all the captions listed in the AutoCaption dialog box. AutoCaption objects cannot be programmatically added to or deleted from the AutoCaptions collection. Use AutoCaptions (_index_), where _index_ is the caption name or index number, to return a single AutoCaption object. The caption names correspond to the items listed in the AutoCaption dialog box. You must exactly match the spelling (but not necessarily the capitalization) of the name, as it is shown in the AutoCaption dialog box. The following example displays the caption text "Microsoft Word Table." The index number represents the position of the AutoCaption object in the list of captions in the AutoCaption dialog box. The following example displays the name of the first item selected in the AutoCaption dialog box.

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Applicationobject that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified AutoCaptions collection.
- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns the number of items in the AutoCaptions collection. Read-only Long.

## Methods (2)

- `Item(Index As Variant) As AutoCaption`  
  Returns an individual AutoCaption object in a collection.
    - `Index As Variant` (required): The individual object to be returned. Can be a Long indicating the ordinal position or a String representing the name of the individual object.
- `CancelAutoInsert()`  
  Prevents Word from automatically adding captions to any type of item.
