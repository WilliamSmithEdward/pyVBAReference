# IVBSAXContentHandler

**Type:** Dispatch Interface  
**Library:** Microsoft XML, v6.0  
**GUID:** {2ED7290A-4DD5-4B46-BB26-4E4155E77FAA}  

IVBSAXContentHandler interface

## Properties (1)

- `documentLocator As IVBSAXLocator  (write-only)`  
  Receive an object for locating the origin of SAX document events.

## Methods (10)

- `startDocument()`  
  Receive notification of the beginning of a document.
- `endDocument()`  
  Receive notification of the end of a document.
- `startPrefixMapping(strPrefix As String, strURI As String)`  
  Begin the scope of a prefix-URI Namespace mapping.
- `endPrefixMapping(strPrefix As String)`  
  End the scope of a prefix-URI mapping.
- `startElement(strNamespaceURI As String, strLocalName As String, strQName As String, oAttributes As IVBSAXAttributes)`  
  Receive notification of the beginning of an element.
- `endElement(strNamespaceURI As String, strLocalName As String, strQName As String)`  
  Receive notification of the end of an element.
- `characters(strChars As String)`  
  Receive notification of character data.
- `ignorableWhitespace(strChars As String)`  
  Receive notification of ignorable whitespace in element content.
- `processingInstruction(strTarget As String, strData As String)`  
  Receive notification of a processing instruction.
- `skippedEntity(strName As String)`  
  Receive notification of a skipped entity.
