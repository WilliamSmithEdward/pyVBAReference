# HeadersFooters

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020984-0000-0000-C000-000000000046}  

A collection of HeaderFooter objects that represent the headers or footers in the specified section of a document.

**Remarks:** Use the Headers or Footers property to return the HeadersFooters collection. The following example displays the text from the primary footer in the first section of the active document. Use Headers (Index) or Footers (Index), where index is one of the WdHeaderFooterIndex constants (wdHeaderFooterEvenPages, wdHeaderFooterFirstPage, or wdHeaderFooterPrimary), to return a single HeaderFooter object. The following example changes the text of both the primary header and the primary footer the first section of the active document. You can also return a single HeaderFooter object by using the HeaderFooter property with a Selection object. Use the DifferentFirstPageHeaderFooter property with the PageSetup object to specify a different first page. The following example inserts text into the first page footer in the active document. Use the OddAndEvenPagesHeaderFooter property with the PageSetup object to specify different odd and even page headers and footers. If the OddAndEvenPagesHeaderFooter property is True, you can return an odd header or footer by using wdHeaderFooterPrimary, and you can return an even header or footer by using wdHeaderFooterEvenPages.

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified HeadersFooters object.
- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of headers and/or footers in the collection. Read-only.

## Methods (1)

- `Item(Index As WdHeaderFooterIndex) As HeaderFooter`  
  Returns a HeaderFooter object that represents a header or footer in a range or section.
    - `Index As WdHeaderFooterIndex` (required): A constant that specifies the header or footer in the range or section.
