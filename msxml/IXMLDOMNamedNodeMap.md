# IXMLDOMNamedNodeMap

**Type:** Dispatch Interface  
**Library:** Microsoft XML, v6.0  
**GUID:** {2933BF83-7B36-11D2-B20E-00C04F983E60}  

## Properties (2)

- `item As IXMLDOMNode  (read-only)`  
  collection of nodes
- `length As Long  (read-only)`  
  number of nodes in the collection

## Methods (7)

- `getNamedItem(name As String) As IXMLDOMNode`  
  lookup item by name
- `setNamedItem(newItem As IXMLDOMNode) As IXMLDOMNode`  
  set item by name
- `removeNamedItem(name As String) As IXMLDOMNode`  
  remove item by name
- `getQualifiedItem(baseName As String, namespaceURI As String) As IXMLDOMNode`  
  lookup the item by name and namespace
- `removeQualifiedItem(baseName As String, namespaceURI As String) As IXMLDOMNode`  
  remove the item by name and namespace
- `nextNode() As IXMLDOMNode`  
  get next node from iterator
- `reset()`  
  reset the position of iterator
