# IVBSAXAttributes

**Type:** Dispatch Interface  
**Library:** Microsoft XML, v6.0  
**GUID:** {10DC0586-132B-4CAC-8BB3-DB00AC8B7EE0}  

IVBSAXAttributes interface

## Properties (1)

- `length As Long  (read-only)`  
  Get the number of attributes in the list.

## Methods (11)

- `getURI(nIndex As Long) As String`  
  Look up an attribute's Namespace URI by index.
- `getLocalName(nIndex As Long) As String`  
  Look up an attribute's local name by index.
- `getQName(nIndex As Long) As String`  
  Look up an attribute's XML 1.0 qualified name by index.
- `getIndexFromName(strURI As String, strLocalName As String) As Long`  
  Look up the index of an attribute by Namespace name.
- `getIndexFromQName(strQName As String) As Long`  
  Look up the index of an attribute by XML 1.0 qualified name.
- `getType(nIndex As Long) As String`  
  Look up an attribute's type by index.
- `getTypeFromName(strURI As String, strLocalName As String) As String`  
  Look up an attribute's type by Namespace name.
- `getTypeFromQName(strQName As String) As String`  
  Look up an attribute's type by XML 1.0 qualified name.
- `getValue(nIndex As Long) As String`  
  Look up an attribute's value by index.
- `getValueFromName(strURI As String, strLocalName As String) As String`  
  Look up an attribute's value by Namespace name.
- `getValueFromQName(strQName As String) As String`  
  Look up an attribute's value by XML 1.0 qualified name.
