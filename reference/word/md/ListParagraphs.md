# ListParagraphs

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020991-0000-0000-C000-000000000046}  

A collection of Paragraph objects that represents the paragraphs of the specified document, list, or range that have list formatting applied.

**Remarks:** Use the ListParagraphs property to return the ListParagraphs collection. The following example applies highlighting to the collection of paragraphs with list formatting in the active document. Use ListParagraphs (Index), where Index is the index number, to return a single Paragraph object with list formatting. Paragraphs can have two types of list formatting. The first type includes an automatically added number or bullet at the beginning of each paragraph in the list. The second type includes LISTNUM fields, which can be placed anywhere inside a paragraph. There can be more than one LISTNUM field per paragraph. To add list formatting to paragraphs, you can use the ApplyListTemplate, ApplyBulletDefault, ApplyNumberDefault, or ApplyOutlineNumberDefault method. You access these methods through the ListFormat object for a specified range. The Count property for this collection in a document returns the number of items in the main story only. To count items in other stories use the collection with the Range object.

## Properties (5)

- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of list paragraphs in the collection. Read-only.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified ListParagraphs object.

## Methods (1)

- `Item(Index As Long) As Paragraph`  
  Returns an individual Paragraph object in a collection.
    - `Index As Long` (required): The individual object to be returned. Can be a Long indicating the ordinal position of the individual object.
