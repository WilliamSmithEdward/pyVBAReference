# HeaderFooter

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020985-0000-0000-C000-000000000046}  

Represents a single header or footer. The HeaderFooter object is a member of the HeadersFooters collection. The HeadersFooters collection includes all headers and footers in the specified document section.

**Remarks:** Use Headers (Index) or Footers (Index), where index is one of the WdHeaderFooterIndex constants (wdHeaderFooterEvenPages, wdHeaderFooterFirstPage, or wdHeaderFooterPrimary), to return a single HeaderFooter object. The following example changes the text of both the primary header and the primary footer in the first section of the active document. You can also return a single HeaderFooter object by using the HeaderFooter property with a Selection object. Use the DifferentFirstPageHeaderFooter property with the PageSetup object to specify a different first page. The following example inserts text into the first page footer in the active document. Use the OddAndEvenPagesHeaderFooter property with the PageSetup object to specify different odd and even page headers and footers. If the OddAndEvenPagesHeaderFooter property is True, you can return an odd header or footer by using wdHeaderFooterPrimary, and you can return an even header or footer by using wdHeaderFooterEvenPages. Use the Add method with the PageNumbers object to add a page number to a header or footer. The following example adds page numbers to the primary footer in the first section of the active document.

## Properties (11)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified HeaderFooter object.
- `Range As Range  (read-only)`  
  Returns a Range object that represents the portion of a document that is contained within the specified header or footer.
- `Index As WdHeaderFooterIndex  (read-only)`  
  Returns a WdHeaderFooterIndex that represents the specified header or footer in a document or section. Read-only.
- `IsHeader As Boolean  (read-only)`  
  True if the specified HeaderFooter object is a header. Read-only Boolean.
- `Exists As Boolean  (read/write)`  
  True if the specified HeaderFooter object exists. Read/write Boolean.
- `PageNumbers As PageNumbers  (read-only)`  
  Returns a PageNumbers collection that represents all the page number fields included in the specified header or footer.
- `LinkToPrevious As Boolean  (read/write)`  
  True if the specified header or footer is linked to the corresponding header or footer in the previous section. Read/write Boolean.
- `Shapes As Shapes  (read-only)`  
  Returns a Shapes collection that represents all the Shape objects in a header or footer. Read-only.
- `IsEmpty As Boolean  (read-only)`
