# ISAXContentHandler

**Type:** Interface  
**Library:** Microsoft XML, v6.0  
**GUID:** {1545CDFA-9E4E-4497-A8A4-2BF7D0112C44}  

ISAXContentHandler interface

## Methods (11)

- `putDocumentLocator(pLocator As ISAXLocator)`
- `startDocument()`
- `endDocument()`
- `startPrefixMapping(pwchPrefix As Integer, cchPrefix As Long, pwchUri As Integer, cchUri As Long)`
- `endPrefixMapping(pwchPrefix As Integer, cchPrefix As Long)`
- `startElement(pwchNamespaceUri As Integer, cchNamespaceUri As Long, pwchLocalName As Integer, cchLocalName As Long, pwchQName As Integer, cchQName As Long, pAttributes As ISAXAttributes)`
- `endElement(pwchNamespaceUri As Integer, cchNamespaceUri As Long, pwchLocalName As Integer, cchLocalName As Long, pwchQName As Integer, cchQName As Long)`
- `characters(pwchChars As Integer, cchChars As Long)`
- `ignorableWhitespace(pwchChars As Integer, cchChars As Long)`
- `processingInstruction(pwchTarget As Integer, cchTarget As Long, pwchData As Integer, cchData As Long)`
- `skippedEntity(pwchName As Integer, cchName As Long)`
