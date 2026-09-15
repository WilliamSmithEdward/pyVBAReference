# ListGalleries

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020995-0000-0000-C000-000000000046}  

A collection of ListGallery objects that represent the three tabs in the Bullets and Numbering dialog box.

**Remarks:** Use the ListGalleries property to return the ListGalleries collection. The following code example enumerates the collection of list galleries and sets each of the seven list templates (formats) back to the list template format built into Word. Use ListGalleries (Index), where Index is wdBulletGallery, wdNumberGallery, or wdOutlineNumberGallery, to return a single ListGallery object. The following code example returns the third list format (excluding None) on the Bulleted tab in the Bullets and Numbering dialog box and then applies it to the selection. To see whether the specified list template contains the formatting built into Word, use the Modified property with the ListGallery object. To reset formatting to the original list format, use the Reset method for the ListGallery object.

## Properties (5)

- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of list galleries in the collection. Read-only.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application. Read-only.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified ListGalleries object.

## Methods (1)

- `Item(Index As WdListGalleryType) As ListGallery`  
  Returns an individual ListGalleries object in a collection.
    - `Index As WdListGalleryType` (required): The individual object to be returned.
