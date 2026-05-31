# IXMLDOMSelection

**Type:** Dispatch Interface  
**Library:** Microsoft XML, v6.0  
**GUID:** {AA634FC7-5888-44A7-A257-3A47150D3A0E}  

## Properties (4)

- `item As IXMLDOMNode  (read-only)`  
  collection of nodes
- `length As Long  (read-only)`  
  number of nodes in the collection
- `expr As String  (read/write)`  
  selection expression
- `context As IXMLDOMNode  (read/write)`  
  nodes to apply selection expression to

## Methods (9)

- `nextNode() As IXMLDOMNode`  
  get next node from iterator
- `reset()`  
  reset the position of iterator
- `peekNode() As IXMLDOMNode`  
  gets the next node without advancing the list position
- `matches(pNode As IXMLDOMNode) As IXMLDOMNode`  
  checks to see if the node matches the pattern
- `removeNext() As IXMLDOMNode`  
  removes the next node
- `removeAll()`  
  removes all the nodes that match the selection
- `clone() As IXMLDOMSelection`  
  clone this object with the same position and context
- `getProperty(name As String) As Variant`  
  get the value of the named property
- `setProperty(name As String, value As Variant)`  
  set the value of the named property
