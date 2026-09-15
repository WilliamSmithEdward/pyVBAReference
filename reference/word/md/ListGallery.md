# ListGallery

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020994-0000-0000-C000-000000000046}  

Represents a single gallery of list formats. The ListGallery object is a member of the ListGalleries collection.

**Remarks:** Each ListGallery object represents one of the three tabs in the Bullets and Numbering dialog box. Use ListGalleries (Index), where Index is wdBulletGallery, wdNumberGallery, or wdOutlineNumberGallery, to return a single ListGallery object. The following example returns the third list format (excluding None) on the Bulleted tab in the Bullets and Numbering dialog box and then applies it to the selection. To see whether the specified list template contains the formatting built into Word, use the Modified property for the ListGallery object. To reset formatting to the original list format, use the Reset method for the ListGallery object.

## Properties (5)

- `ListTemplates As ListTemplates  (read-only)`  
  Returns a ListTemplates collection that represents all the list formats for the specified list gallery. Read-only.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified ListGallery object.
- `Modified As Boolean  (read-only)`  
  True if the specified list template is not the built-in list template for that position in the list gallery. Read-only Boolean.

## Methods (1)

- `Reset(Index As Long)`  
  Resets the list template specified by Index for the specified list gallery to the built-in list template format.
    - `Index As Long` (required): The template to reset.
