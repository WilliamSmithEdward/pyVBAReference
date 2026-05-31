# IVBSAXLexicalHandler

**Type:** Dispatch Interface  
**Library:** Microsoft XML, v6.0  
**GUID:** {032AAC35-8C0E-4D9D-979F-E3B702935576}  

IVBSAXLexicalHandler interface

## Methods (7)

- `startDTD(strName As String, strPublicId As String, strSystemId As String)`  
  Report the start of DTD declarations, if any.
- `endDTD()`  
  Report the end of DTD declarations.
- `startEntity(strName As String)`  
  Report the beginning of some internal and external XML entities.
- `endEntity(strName As String)`  
  Report the end of an entity.
- `startCDATA()`  
  Report the start of a CDATA section.
- `endCDATA()`  
  Report the end of a CDATA section.
- `comment(strChars As String)`  
  Report an XML comment anywhere in the document.
