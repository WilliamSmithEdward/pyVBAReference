# AutoCaption

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {0002097B-0000-0000-C000-000000000046}  

Represents a single caption that can be automatically added when items such as tables, pictures, or OLE objects are inserted into a document. The AutoCaption object is a member of the AutoCaptions collection. The AutoCaptions collection contains all the captions listed in the AutoCaption dialog box.

**Remarks:** Use AutoCaptions (_index_), where _index_ is the caption name or index number, to return a single AutoCaption object. The caption names correspond to the items listed in the AutoCaption dialog box. You must exactly match the spelling (but not necessarily the capitalization) of the name, as it is shown in the AutoCaption dialog box. The following example enables autocaptions for Word tables. The index number represents the position of the AutoCaption object in the list of items in the AutoCaption dialog box. The following example displays the name of the first item listed in the AutoCaption dialog box. AutoCaption objects cannot be programmatically added to or deleted from the AutoCaptions collection.

## Properties (7)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified AutoCaption object.
- `Name As String  (read-only)`  
  Returns or sets the name of the specified object. Read-only String.
- `AutoInsert As Boolean  (read/write)`  
  True if a caption is automatically added when the item is inserted into a document. Read/write Boolean.
- `Index As Long  (read-only)`  
  Returns a Long that represents the position of an item in a collection. Read-only.
- `CaptionLabel As Variant  (read/write)`  
  Returns or sets the caption label ("Figure," "Table," or "Equation," for example) of the specified caption. Read/write Variant.
