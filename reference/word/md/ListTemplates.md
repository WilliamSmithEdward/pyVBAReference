# ListTemplates

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020990-0000-0000-C000-000000000046}  

A collection of ListTemplate objects in a document, list gallery, or template.

**Remarks:** Use the ListTemplates property with a Document, ListGallery, or Template object to return a ListTemplates collection. With a ListGallery object, the ListTemplates collection is the seven list formats for bulleted lists, numbered lists, and outline numbered lists. The following example displays a message with the level status (single or multiple-level) for each list template in the active document. Use the Add method to add a list template to the collection in the specified document or template. The following example adds a new list template to the active document and applies it to the selection. Use ListTemplates (Index), where Index is the name of a list template or an index number, to return a single list template in a document or template. The following example sets an object variable equal to a list template named "ListBullets" in the active document, and then formats the selection as the first level of that list template. Use ListTemplates (Index), where Index is a number 1 through 7, to return a single list template in a list gallery.

## Properties (5)

- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of list templates in the collection. Read-only.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified ListTemplates object.

## Methods (2)

- `Item(Index As Variant) As ListTemplate`  
  Returns an individual ListTemplate object in a collection.
    - `Index As Variant` (required): The individual object to be returned. Can be a Long indicating the ordinal position or a String representing the name of the individual object.
- `Add([OutlineNumbered As Variant], [Name As Variant]) As ListTemplate`  
  Returns a ListTemplate object that represents a new list template.
    - `OutlineNumbered As Variant` (optional): True to apply outline numbering to the new list template.
    - `Name As Variant` (optional): An optional name used for linking the list template to a LISTNUM field. Use this name to index the list template in the collection.
