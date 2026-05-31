# CustomXMLPart

**Type:** Class  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000CDB08-0000-0000-C000-000000000046}  

Represents a single CustomXMLPart in a CustomXMLParts collection.

**Example:**

```vba
Sub AddPartToCollection()
    Dim myPart As CustomXMLPart

    Set myPart = ActiveDocument.CustomXMLParts.Add("<author>Mark Twain</author>")

End Sub
```

## Properties (11)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the CustomXMLPart object. Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the CustomXMLPart object was created. Read-only.
- `Parent As Object  (read-only)`  
  Gets the Parent object for the CustomXMLPart object. Read-only.
- `DocumentElement As CustomXMLNode  (read-only)`  
  Gets the root element of a bound region of data in a document. If the region is empty, the property returns Nothing. Read-only.
- `Id As String  (read-only)`  
  Gets a String containing the GUID assigned to the current CustomXMLPart object. Read-only.
- `NamespaceURI As String  (read-only)`  
  Gets the unique address identifier for the namespace of the CustomXMLPart object. Read-only.
- `SchemaCollection As CustomXMLSchemaCollection  (read/write)`  
  Gets or sets a CustomXMLSchemaCollection object representing the set of schemas attached to a bound region of data in a document. Read/write.
- `NamespaceManager As CustomXMLPrefixMappings  (read-only)`  
  Gets the set of namespace prefix mappings used against the current CustomXMLPart object. Read-only.
- `XML As String  (read-only)`  
  Gets the XML representation of the current CustomXMLPart object. Read-only.
- `Errors As CustomXMLValidationErrors  (read-only)`  
  Gets a CustomXMLValidationErrors object that provides access to any XML validation errors, if any exist. If no validation errors exist, this property returns Nothing. Read-only.
- `BuiltIn As Boolean  (read-only)`  
  Gets a value that indicates whether the CustomXMLPart is built-in. Read-only.

## Methods (6)

- `AddNode(Parent As CustomXMLNode, [Name As String], [NamespaceURI As String], [NextSibling As CustomXMLNode], [NodeType As MsoCustomXMLNodeType], [NodeValue As String])`  
  Adds a node to the XML tree.
    - `Parent As CustomXMLNode` (required): Represents the node under which this node should be added. If adding an attribute, the parameter denotes the element that the attribute should be added to.
    - `Name As String` (optional): Represents the base name of the node to be added.
    - `NamespaceURI As String` (optional): Represents the namespace of the element to be appended. This parameter is required to append nodes of type msoCustomXMLNodeElement or msoCustomXMLNodeAttribute; otherwise, it is ignored.
    - `NextSibling As CustomXMLNode` (optional): Represents the node which should become the next sibling of the new node. If not specified, the node is added to the end of the parent node's children. This parameter is ignored for additions of type msoXMLNodeAttribute. If the node is not a child of the parent, an error is displayed.
    - `NodeType As MsoCustomXMLNodeType` (optional): Specifies the type of node to append. If the parameter is not specified, it is assumed to be of type msoCustomXMLNodeElement.
    - `NodeValue As String` (optional): Used to set the value of the appended node for those nodes that allow text. If the node doesn't allow text, the parameter is ignored.
- `Delete()`  
  Deletes the current CustomXMLPart from the data store (IXMLDataStore interface).
- `Load(FilePath As String) As Boolean`  
  Allows the template author to populate a CustomXMLPart from an existing file. Returns True if the load was successful.
    - `FilePath As String` (required): Points to the file on the user's computer or on a network containing the XML to be loaded.
- `LoadXML(XML As String) As Boolean`  
  Allows the template author to populate a CustomXMLPart object from an XML string. Returns True if the load was successful.
    - `XML As String` (required): Contains the XML to load.
- `SelectNodes(XPath As String) As CustomXMLNodes`  
  Selects a collection of nodes from a custom XML part.
    - `XPath As String` (required): Contains the XPath expression.
- `SelectSingleNode(XPath As String) As CustomXMLNode`  
  Selects a single node within a custom XML part matching an XPath expression.
    - `XPath As String` (required): Contains an XPath expression.

## Events (3)

- `NodeAfterInsert(NewNode As CustomXMLNode, InUndoRedo As Boolean)`  
  Occurs after a node is inserted in a CustomXMLPart object.
    - `NewNode As CustomXMLNode` (required): Corresponds to the node just added to the CustomXMLPart object. Note that this node may have children if a subtree was just added to the document.
    - `InUndoRedo As Boolean` (required): Returns True if the node was inserted as part of an Undo/Redo action by the user.
- `NodeAfterDelete(OldNode As CustomXMLNode, OldParentNode As CustomXMLNode, OldNextSibling As CustomXMLNode, InUndoRedo As Boolean)`  
  Occurs after a node is deleted in a CustomXMLPart object.
    - `OldNode As CustomXMLNode` (required): Corresponds to the node that was just removed from the CustomXMLPart object. Note that this node may have children if a subtree is being removed from the document. Also, this node will be a "disconnected" node in that you can query down from the node, but you cannot query up the tree; the node appears to exist alone.
    - `OldParentNode As CustomXMLNode` (required): Corresponds to the former parent node of _OldNode_.
    - `OldNextSibling As CustomXMLNode` (required): Corresponds to the former next sibling of _OldNode_.
    - `InUndoRedo As Boolean` (required): Returns True if the node was inserted as part of an Undo/Redo action by the user.
- `NodeAfterReplace(OldNode As CustomXMLNode, NewNode As CustomXMLNode, InUndoRedo As Boolean)`  
  Occurs just after a node is replaced in a CustomXMLPart object.
    - `OldNode As CustomXMLNode` (required): Corresponds to the node that was just removed from the CustomXMLPart object. Note that this node may have children if a subtree was just added to the document. Also, this node will be a "disconnected" node in that you can query down from the node, but cannot go up; it appears to exist alone.
    - `NewNode As CustomXMLNode` (required): Corresponds to the node just added to the CustomXMLPart object.
    - `InUndoRedo As Boolean` (required): Returns True if the node was added as part of an Undo/Redo action by the user.
