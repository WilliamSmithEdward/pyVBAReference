# IVBSAXDeclHandler

**Type:** Dispatch Interface  
**Library:** Microsoft XML, v6.0  
**GUID:** {E8917260-7579-4BE1-B5DD-7AFBFA6F077B}  

IVBSAXDeclHandler interface

## Methods (4)

- `elementDecl(strName As String, strModel As String)`  
  Report an element type declaration.
- `attributeDecl(strElementName As String, strAttributeName As String, strType As String, strValueDefault As String, strValue As String)`  
  Report an attribute type declaration.
- `internalEntityDecl(strName As String, strValue As String)`  
  Report an internal entity declaration.
- `externalEntityDecl(strName As String, strPublicId As String, strSystemId As String)`  
  Report a parsed external entity declaration.
