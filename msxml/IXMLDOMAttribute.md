# IXMLDOMAttribute

**Type:** Dispatch Interface  
**Library:** Microsoft XML, v6.0  
**GUID:** {2933BF85-7B36-11D2-B20E-00C04F983E60}  

## Properties (24)

- `nodeName As String  (read-only)`  
  name of the node
- `nodeValue As Variant  (read/write)`  
  value stored in the node
- `nodeType As DOMNodeType  (read-only)`  
  the node's type
- `parentNode As IXMLDOMNode  (read-only)`  
  parent of the node
- `childNodes As IXMLDOMNodeList  (read-only)`  
  the collection of the node's children
- `firstChild As IXMLDOMNode  (read-only)`  
  first child of the node
- `lastChild As IXMLDOMNode  (read-only)`  
  last child of the node
- `previousSibling As IXMLDOMNode  (read-only)`  
  left sibling of the node
- `nextSibling As IXMLDOMNode  (read-only)`  
  right sibling of the node
- `attributes As IXMLDOMNamedNodeMap  (read-only)`  
  the collection of the node's attributes
- `ownerDocument As IXMLDOMDocument  (read-only)`  
  document that contains the node
- `nodeTypeString As String  (read-only)`  
  the type of node in string form
- `text As String  (read/write)`  
  text content of the node and subtree
- `specified As Boolean  (read-only)`  
  indicates whether node is a default value
- `definition As IXMLDOMNode  (read-only)`  
  pointer to the definition of the node in the DTD or schema
- `nodeTypedValue As Variant  (read/write)`  
  get the strongly typed value of the node
- `dataType As Variant  (read/write)`  
  the data type of the node
- `xml As String  (read-only)`  
  return the XML source for the node and each of its descendants
- `parsed As Boolean  (read-only)`  
  has sub-tree been completely parsed
- `namespaceURI As String  (read-only)`  
  the URI for the namespace applying to the node
- `prefix As String  (read-only)`  
  the prefix for the namespace applying to the node
- `baseName As String  (read-only)`  
  the base name of the node (nodename with the prefix stripped off)
- `name As String  (read-only)`  
  get name of the attribute
- `value As Variant  (read/write)`  
  string value of the attribute

## Methods (10)

- `insertBefore(newChild As IXMLDOMNode, refChild As Variant) As IXMLDOMNode`  
  insert a child node
- `replaceChild(newChild As IXMLDOMNode, oldChild As IXMLDOMNode) As IXMLDOMNode`  
  replace a child node
- `removeChild(childNode As IXMLDOMNode) As IXMLDOMNode`  
  remove a child node
- `appendChild(newChild As IXMLDOMNode) As IXMLDOMNode`  
  append a child node
- `hasChildNodes() As Boolean`
- `cloneNode(deep As Boolean) As IXMLDOMNode`
- `transformNode(stylesheet As IXMLDOMNode) As String`  
  apply the stylesheet to the subtree
- `selectNodes(queryString As String) As IXMLDOMNodeList`  
  execute query on the subtree
- `selectSingleNode(queryString As String) As IXMLDOMNode`  
  execute query on the subtree
- `transformNodeToObject(stylesheet As IXMLDOMNode, outputObject As Variant)`  
  apply the stylesheet to the subtree, returning the result through a document or a stream
