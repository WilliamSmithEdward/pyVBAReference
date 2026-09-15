# XMLNode

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {09760240-0B89-49F7-A79D-479F24723F56}  

Represents a single XML element applied to a document.

**Remarks:** Each XML element that has been applied to a document is displayed as a node in a tree view control in the XML Structure task pane. Each node in the tree view is an instance of an XMLNode object. The hierarchy in the tree view indicates whether a node contains child nodes. Use the Item method of the XMLNodes collection to return an individual XMLNode object. Use the Validate method to verify that an XML element is valid according to the applied schemas and that any required child elements exist and are in the required order. Once you run the Validate method, use the ValidationStatus property to verify whether an element is valid, and use the ValidationErrorText property to display information about what the user needs to do to make the document conform to the XML schema rules. The following example validates each of the XML elements in the active document. If the element is found to be invalid against the schema, the example returns a message to the user explaining what the problem is.

## Properties (24)

- `BaseName As String  (read-only)`  
  Returns a String that represents the name of the element without any prefix.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified XMLNode object.
- `Range As Range  (read-only)`  
  Returns a Range object that represents the portion of a document that is contained in the specified object. Read-only.
- `Text As String  (read/write)`  
  Returns or sets the text contained within the XML element. Read/write String.
- `NamespaceURI As String  (read-only)`  
  Returns a String that represents the Uniform Resource Identifier (URI) of the schema namespace for the specified object. Read-only.
- `XML As String  (read-only)`
- `NextSibling As XMLNode  (read-only)`  
  Returns an XMLNode object that represents the next element in the document that is at the same level as the specified element.
- `PreviousSibling As XMLNode  (read-only)`  
  Returns an XMLNode object that represents the previous element in the document that is at the same level as the specified element.
- `ParentNode As XMLNode  (read-only)`  
  Returns an XMLNode object that represents the parent element of the specified element.
- `FirstChild As XMLNode  (read-only)`  
  Returns a DiagramNode object that represents the first child node of a parent node. Read-only.
- `LastChild As XMLNode  (read-only)`  
  Returns an XMLNode object that represents the last child node of an XML element.
- `OwnerDocument As Document  (read-only)`  
  Returns a Document object that represents the parent document of the specified XML element.
- `NodeType As WdXMLNodeType  (read-only)`  
  Returns a WdXMLNodeType constant that represents the type of node.
- `ChildNodes As XMLNodes  (read-only)`  
  Returns an XMLNodes collection that represents the child elements of a specified element.
- `Attributes As XMLNodes  (read-only)`  
  Returns an XMLNodes collection that represents the attributes for the specified element.
- `NodeValue As String  (read/write)`  
  Returns or sets a String that represents the value of an XML element. Read/write.
- `HasChildNodes As Boolean  (read-only)`  
  Returns a Boolean that represents whether an XML node has child nodes. Read-only.
- `Level As WdXMLNodeLevel  (read-only)`  
  Returns a WdXMLNodeLevel constant that represents whether an XML element is part of a paragraph, is a paragraph, or is contained within a table cell or contains a table row. Read-only.
- `ValidationStatus As WdXMLValidationStatus  (read-only)`  
  Returns a WdXMLValidationStatus constant that represents whether an element or attribute is valid according to the attached schema.
- `ValidationErrorText As String  (read-only)`  
  Returns a String that represents the description for a validation error on an XMLNode object.
- `PlaceholderText As String  (read/write)`  
  Sets or returns a String that represents the text displayed for an element that contains no text.
- `WordOpenXML As String  (read-only)`  
  Returns a String that represents the XML for the node in the Microsoft Word Open XML format. Read-only.

## Methods (8)

- `SelectSingleNode(XPath As String, [PrefixMapping As String], [FastSearchSkippingTextNodes As Boolean]) As XMLNode`  
  Returns an XMLNode object that represents the first child element that matches the XPath parameter within the specified XML element. .
    - `XPath As String` (required): Specifies a valid XPath string. For more information on XPath, see the XPath reference documentation on the Microsoft Developer Network (MSDN) Web site.
    - `PrefixMapping As String` (optional): Provides the prefix in the schema against which to perform the search. Use the PrefixMapping parameter if your XPath parameter uses names to search for elements.
    - `FastSearchSkippingTextNodes As Boolean` (optional): True skips all text nodes while searching for the specified node. False includes text nodes in the search. Default value is False.
- `SelectNodes(XPath As String, [PrefixMapping As String], [FastSearchSkippingTextNodes As Boolean]) As XMLNodes`  
  Returns an XMLNodes collection that represents all the child elements that match the XPath parameter, in the order in which they appear within the specified XML element.
    - `XPath As String` (required): Specifies a valid XPath string. For more information on XPath, see the XPath reference documentation on the Microsoft Developer Network (MSDN) Web site.
    - `PrefixMapping As String` (optional): Provides the prefix in the schema against which to perform the search. Use the PrefixMapping parameter if your XPath parameter uses names to search for elements.
    - `FastSearchSkippingTextNodes As Boolean` (optional): True skips all text nodes while searching for the specified node. False includes text nodes in the search. Default value is False.
- `Delete()`  
  Deletes the specified XML element from an XML document.
- `Copy()`  
  Copies the specified XML element, excluding XML markup, to the Clipboard.
- `RemoveChild(ChildElement As XMLNode)`  
  Removes a child element from the specified element.
    - `ChildElement As XMLNode` (required): The child element to be removed.
- `Cut()`  
  Removes the specified XML element from the document and places it on the Clipboard.
- `Validate()`  
  Validates an individual XML element against the XML schemas that are attached to a document.
- `SetValidationError(Status As WdXMLValidationStatus, [ErrorText As Variant], [ClearedAutomatically As Boolean])`  
  Changes the validation error text displayed to a user for a specified node and forces Word to report a node as invalid.
    - `Status As WdXMLValidationStatus` (required): Specifies whether to set the validation status error text (wdXMLValidationStatusCustom) or to clear the validation status error text (wdXMLValidationStatusOK).
    - `ErrorText As Variant` (optional): The text displayed to the user. Leave blank when the Status parameter is set to wdXMLValidationStatusOK.
    - `ClearedAutomatically As Boolean` (optional): True automatically clears the error message as soon as the next validation event occurs on the specified node. False requires running the SetValidationError method with a Status parameter of wdXMLValidationStatusOK to clear the custom error text.
