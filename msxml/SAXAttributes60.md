# SAXAttributes60

**Type:** Class  
**Library:** Microsoft XML, v6.0  
**GUID:** {88D96A0E-F192-11D4-A65F-0040963251E5}  

SAX Attributes 6.0

## Methods (11)

- `addAttribute(strURI As String, strLocalName As String, strQName As String, strType As String, strValue As String)`  
  Add an attribute to the end of the list.
- `addAttributeFromIndex(varAtts As Variant, nIndex As Long)`  
  Add an attribute, whose value is equal to the indexed attribute in the input attributes object, to the end of the list.
- `clear()`  
  Clear the attribute list for reuse.
- `removeAttribute(nIndex As Long)`  
  Remove an attribute from the list.
- `setAttribute(nIndex As Long, strURI As String, strLocalName As String, strQName As String, strType As String, strValue As String)`  
  Set an attribute in the list.
- `setAttributes(varAtts As Variant)`  
  Copy an entire Attributes object.
- `setLocalName(nIndex As Long, strLocalName As String)`  
  Set the local name of a specific attribute.
- `setQName(nIndex As Long, strQName As String)`  
  Set the qualified name of a specific attribute.
- `setType(nIndex As Long, strType As String)`  
  Set the type of a specific attribute.
- `setURI(nIndex As Long, strURI As String)`  
  Set the Namespace URI of a specific attribute.
- `setValue(nIndex As Long, strValue As String)`  
  Set the value of a specific attribute.
