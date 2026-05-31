# DOMDocument60

**Type:** Class  
**Library:** Microsoft XML, v6.0  
**GUID:** {88D96A05-F192-11D4-A65F-0040963251E5}  

W3C-DOM XML Document 6.0 (Apartment)

## Properties (37)

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
- `doctype As IXMLDOMDocumentType  (read-only)`  
  node corresponding to the DOCTYPE
- `implementation As IXMLDOMImplementation  (read-only)`  
  info on this DOM implementation
- `documentElement As IXMLDOMElement  (read/write)`  
  the root of the tree
- `readyState As Long  (read-only)`  
  get the state of the XML document
- `parseError As IXMLDOMParseError  (read-only)`  
  get the last parser error
- `url As String  (read-only)`  
  get the URL for the loaded XML document
- `async As Boolean  (read/write)`  
  flag for asynchronous download
- `validateOnParse As Boolean  (read/write)`  
  indicates whether the parser performs validation
- `resolveExternals As Boolean  (read/write)`  
  indicates whether the parser resolves references to external DTD/Entities/Schema
- `preserveWhiteSpace As Boolean  (read/write)`  
  indicates whether the parser preserves whitespace
- `onreadystatechange As Variant  (write-only)`  
  register a readystatechange event handler
- `ondataavailable As Variant  (write-only)`  
  register an ondataavailable event handler
- `ontransformnode As Variant  (write-only)`  
  register an ontransformnode event handler
- `namespaces As IXMLDOMSchemaCollection  (read-only)`  
  A collection of all namespaces for this document
- `schemas As Variant  (read/write)`  
  The associated schema cache

## Methods (30)

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
- `createElement(tagName As String) As IXMLDOMElement`  
  create an Element node
- `createDocumentFragment() As IXMLDOMDocumentFragment`  
  create a DocumentFragment node
- `createTextNode(data As String) As IXMLDOMText`  
  create a text node
- `createComment(data As String) As IXMLDOMComment`  
  create a comment node
- `createCDATASection(data As String) As IXMLDOMCDATASection`  
  create a CDATA section node
- `createProcessingInstruction(target As String, data As String) As IXMLDOMProcessingInstruction`  
  create a processing instruction node
- `createAttribute(name As String) As IXMLDOMAttribute`  
  create an attribute node
- `createEntityReference(name As String) As IXMLDOMEntityReference`  
  create an entity reference node
- `getElementsByTagName(tagName As String) As IXMLDOMNodeList`  
  build a list of elements by name
- `createNode(type As Variant, name As String, namespaceURI As String) As IXMLDOMNode`  
  create a node of the specified node type and name
- `nodeFromID(idString As String) As IXMLDOMNode`  
  retrieve node from it's ID
- `load(xmlSource As Variant) As Boolean`  
  load document from the specified XML source
- `abort()`  
  abort an asynchronous download
- `loadXML(bstrXML As String) As Boolean`  
  load the document from a string
- `save(destination As Variant)`  
  save the document to a specified destination
- `validate() As IXMLDOMParseError`  
  perform runtime validation on the currently loaded XML document
- `setProperty(name As String, value As Variant)`  
  set the value of the named property
- `getProperty(name As String) As Variant`  
  get the value of the named property
- `validateNode(node As IXMLDOMNode) As IXMLDOMParseError`  
  perform runtime validation on the currently loaded XML document node
- `importNode(node As IXMLDOMNode, deep As Boolean) As IXMLDOMNode`  
  clone node such that clones ownerDocument is this document

## Events (2)

- `ondataavailable()`
- `onreadystatechange()`
