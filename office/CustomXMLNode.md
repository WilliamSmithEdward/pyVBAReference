# CustomXMLNode

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000CDB04-0000-0000-C000-000000000046}  

## Properties (19)

- `Application As Object  (read-only)`
- `Creator As Long  (read-only)`
- `Parent As Object  (read-only)`
- `Attributes As CustomXMLNodes  (read-only)`
- `BaseName As String  (read-only)`
- `ChildNodes As CustomXMLNodes  (read-only)`
- `FirstChild As CustomXMLNode  (read-only)`
- `LastChild As CustomXMLNode  (read-only)`
- `NamespaceURI As String  (read-only)`
- `NextSibling As CustomXMLNode  (read-only)`
- `NodeType As MsoCustomXMLNodeType  (read-only)`
- `NodeValue As String  (read/write)`
- `OwnerDocument As Object  (read-only)`
- `OwnerPart As CustomXMLPart  (read-only)`
- `PreviousSibling As CustomXMLNode  (read-only)`
- `ParentNode As CustomXMLNode  (read-only)`
- `Text As String  (read/write)`
- `XPath As String  (read-only)`
- `XML As String  (read-only)`

## Methods (11)

- `AppendChildNode([Name As String], [NamespaceURI As String], [NodeType As MsoCustomXMLNodeType], [NodeValue As String])`
- `AppendChildSubtree(XML As String)`
- `Delete()`
- `HasChildNodes() As Boolean`
- `InsertNodeBefore([Name As String], [NamespaceURI As String], [NodeType As MsoCustomXMLNodeType], [NodeValue As String], [NextSibling As CustomXMLNode])`
- `InsertSubtreeBefore(XML As String, [NextSibling As CustomXMLNode])`
- `RemoveChild(Child As CustomXMLNode)`
- `ReplaceChildNode(OldNode As CustomXMLNode, [Name As String], [NamespaceURI As String], [NodeType As MsoCustomXMLNodeType], [NodeValue As String])`
- `ReplaceChildSubtree(XML As String, OldNode As CustomXMLNode)`
- `SelectNodes(XPath As String) As CustomXMLNodes`
- `SelectSingleNode(XPath As String) As CustomXMLNode`
