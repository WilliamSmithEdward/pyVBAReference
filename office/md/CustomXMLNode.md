# CustomXMLNode

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000CDB04-0000-0000-C000-000000000046}  

Represents an XML node in a tree in a document. The CustomXMLNode object is a member of the CustomXMLNodes collection.

**Remarks:** The CustomXMLNode object is designed to have functional parity with the IXMLDOMNode interface. In addition, it contains an XPath property, which is a great improvement over the objects provided by MSXML.

**Example:**

```vba
Sub CustomXmlNodes()
    Dim cxp1 As CustomXMLPart
    Dim cxn As CustomXMLNode

    With ActiveDocument

        ' Returns the first custom xml part with the given root namespace.
        Set cxp1 = .CustomXMLParts("urn:invoice:namespace")

        ' Get the first node matching the XPath expression.
        Set cxn = cxp1.SelectSingleNode("//*[@quantity < 4]")

    End With

End Sub
```

## Properties (19)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for a CustomXMLNode object. Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the CustomXMLNode object was created. Read-only.
- `Parent As Object  (read-only)`  
  Gets the Parent object for the CustomXMLNode object. Read-only.
- `Attributes As CustomXMLNodes  (read-only)`  
  Gets a CustomXMLNodes collection representing the attributes of the current element in the current node. Read-only.
- `BaseName As String  (read-only)`  
  Gets the base name of the node without the namespace prefix, if one exists, in the Document Object Model (DOM). Read-only.
- `ChildNodes As CustomXMLNodes  (read-only)`  
  Gets a CustomXMLNodes collection containing all of the child elements of the current node. Read-only.
- `FirstChild As CustomXMLNode  (read-only)`  
  Gets a CustomXMLNode object corresponding to the first child element of the current node. If the node has no child elements (or if it isn't of type msoCustomXMLNodeElement), returns Nothing. Read-only.
- `LastChild As CustomXMLNode  (read-only)`  
  Gets a CustomXMLNode object corresponding to the last child element of the current node. If the node has no child elements (or if it is not of type msoCustomXMLNodeElement), the property returns Nothing. Read-only.
- `NamespaceURI As String  (read-only)`  
  Gets the unique address identifier for the namespace of the CustomXMLNode object. Read-only.
- `NextSibling As CustomXMLNode  (read-only)`  
  Gets the next sibling node (element, comment, or processing instruction) of the current node. If the node is the last sibling at its level, the property returns Nothing. Read-only.
- `NodeType As MsoCustomXMLNodeType  (read-only)`  
  Gets the type of the current node. Read-only.
- `NodeValue As String  (read/write)`  
  Gets or sets the value of the current node. Read/write.
- `OwnerDocument As Object  (read-only)`  
  Gets the object representing the Microsoft Excel workbook, PowerPoint presentation, or Word document associated with this node. Read-only.
- `OwnerPart As CustomXMLPart  (read-only)`  
  Gets the object representing the part associated with this node. Read-only.
- `PreviousSibling As CustomXMLNode  (read-only)`  
  Gets the previous sibling node (element, comment, or processing instruction) of the current node. If the current node is the first sibling at its level, the property returns Nothing. Read-only.
- `ParentNode As CustomXMLNode  (read-only)`  
  Gets the parent element node of the current node. If the current node is at the root level, the property returns Nothing. Read-only.
- `Text As String  (read/write)`  
  Gets or sets the text for the current node. Read/write.
- `XPath As String  (read-only)`  
  Gets a String with the canonicalized XPath for the current node. If the node is no longer in the Document Object Model (DOM), the property returns an error message. Read-only.
- `XML As String  (read-only)`  
  Gets the XML representation of the current node and its children, if any exist. Read-only.

## Methods (11)

- `AppendChildNode([Name As String], [NamespaceURI As String], [NodeType As MsoCustomXMLNodeType], [NodeValue As String])`  
  Appends a single node as the last child under the context element node in the tree.
    - `Name As String` (optional): Represents the base name of the element to be appended.
    - `NamespaceURI As String` (optional): Represents the namespace of the element to be appended. This parameter is required to append nodes of type msoCustomXMLNodeElement or msoCustomXMLNodeAttribute; otherwise, it is ignored.
    - `NodeType As MsoCustomXMLNodeType` (optional): Specifies the type of node to append. If the parameter is not specified, it is assumed to be of type msoCustomXMLNodeElement.
    - `NodeValue As String` (optional): Used to set the value of the appended node for those nodes that allow text. If the node doesn't allow text, the parameter is ignored.
- `AppendChildSubtree(XML As String)`  
  Adds a subtree as the last child under the context element node in the tree.
    - `XML As String` (required): Represents the subtree to add.
- `Delete()`  
  Deletes the current node from the tree (including all of its children, if any exist).
- `HasChildNodes() As Boolean`  
  Gets True if the current element node has child element nodes.
- `InsertNodeBefore([Name As String], [NamespaceURI As String], [NodeType As MsoCustomXMLNodeType], [NodeValue As String], [NextSibling As CustomXMLNode])`  
  Inserts a new node just before the context node in the tree.
    - `Name As String` (optional): Represents the base name of the node to be added.
    - `NamespaceURI As String` (optional): Represents the namespace of the element to be added. This parameter is required if adding nodes of type msoCustomXMLNodeElement or msoCustomXMLNodeAttribute; otherwise, it is ignored.
    - `NodeType As MsoCustomXMLNodeType` (optional): Specifies the type of the node to be added. If the parameter is not specified, it is assumed to be a node of type msoCustomXMLNodeElement.
    - `NodeValue As String` (optional): Used to set the value of the node to be added for those nodes that allow text. If the node doesn't allow text, the parameter is ignored.
    - `NextSibling As CustomXMLNode` (optional): Represents the context node.
- `InsertSubtreeBefore(XML As String, [NextSibling As CustomXMLNode])`  
  Inserts the specified subtree into the location just before the context node.
    - `XML As String` (required): Represents the subtree to be added.
    - `NextSibling As CustomXMLNode` (optional): Specifies the context node.
- `RemoveChild(Child As CustomXMLNode)`  
  Removes the specified child node from the tree.
    - `Child As CustomXMLNode` (required): Represents the child node of the context node.
- `ReplaceChildNode(OldNode As CustomXMLNode, [Name As String], [NamespaceURI As String], [NodeType As MsoCustomXMLNodeType], [NodeValue As String])`  
  Removes the specified child node (and its subtree) from the main tree, and replaces it with a different node in the same location.
    - `OldNode As CustomXMLNode` (required): Represents the child node to be replaced.
    - `Name As String` (optional): Represents the base name of the element to be added.
    - `NamespaceURI As String` (optional): Represents the namespace of the element to be added. This parameter is required if adding nodes of type msoCustomXMLNodeElement or msoCustomXMLNodeAttribute; otherwise, it is ignored.
    - `NodeType As MsoCustomXMLNodeType` (optional): Specifies the type of node to add. If the parameter is not specified, it is assumed to be of type msoCustomXMLNodeElement.
    - `NodeValue As String` (optional): Used to set the value of the node to be added for those nodes that allow text. If the node doesn't allow text, the parameter is ignored.
- `ReplaceChildSubtree(XML As String, OldNode As CustomXMLNode)`  
  Removes the specified node (and its subtree) from the main tree, and replaces it with a different subtree in the same location.
    - `XML As String` (required): Represents the subtree to be added.
    - `OldNode As CustomXMLNode` (required): Represents the child node to be replaced.
- `SelectNodes(XPath As String) As CustomXMLNodes`  
  Selects a collection of nodes matching an XPath expression. This method differs from the CustomXMLPart.SelectNodes method in that the XPath expression will be evaluated starting with the 'expression' node as the context node.
    - `XPath As String` (required): Contains an XPath expression.
- `SelectSingleNode(XPath As String) As CustomXMLNode`  
  Selects a single node from a collection matching an XPath expression. This method differs from the CustomXMLPart.SelectSingleNode method in that the XPath expression will be evaluated starting with the 'expression' node as the context node.
    - `XPath As String` (required): Contains an XPath expression.
