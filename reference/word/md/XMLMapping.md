# XMLMapping

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {0C1FABE7-F737-406F-9CA3-B07661F9D1A2}  

Represents the XML mapping on a ContentControl object between custom XML and a content control. An XML mapping is a link between the text in a content control and an XML element in the custom XML data store for this document.

**Remarks:** Use the SetMapping method to add or change the XML mapping for a content control using an XPath string. The following example sets the built-in document property for the document author, inserts a new content control into the active document, and then sets the XML mapping for the control to the built-in document property. Use the SetMappingByNode method to add or change the XML mapping for a content control using a CustomXMLNode object. The following example does the same thing as the previous example, but uses the SetMappingByNode method. The following example creates a new CustomXMLPart object, loads custom XML into it, and then creates two new content controls and maps each to a different XML element within the custom XML. Use the Delete method to remove the XML mapping for a content control. Deleting the XML mapping for a content control deletes only the connection between the content control and the XML data. Both the content control and the XML data remain in the document. The following example deletes the XML mapping for all content controls in the active document that are currently mapped.

## Properties (8)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified XMLMapping object.
- `IsMapped As Boolean  (read-only)`  
  Returns a Boolean that represents whether the content control in the document is mapped to an XML node in the document's XML data store. Read-only.
- `CustomXMLPart As CustomXMLPart  (read-only)`  
  Returns a CustomXMLPart object that represents the custom XML part to which the content control in the document maps.
- `CustomXMLNode As CustomXMLNode  (read-only)`  
  Returns a CustomXMLNode object that represents the custom XML node in the data store to which the content control in the document maps.
- `XPath As String  (read-only)`  
  Returns a String that represents the XPath for the XML mapping, which evaluates to the currently mapped XML node. Read-only.
- `PrefixMappings As String  (read-only)`  
  Returns a String that represents the prefix mappings used to evaluate the XPath for the current XML mapping. Read-only.

## Methods (3)

- `SetMapping(XPath As String, [PrefixMapping As String], [Source As CustomXMLPart]) As Boolean`  
  Allows creating or changing the XML mapping on a content control. Returns True if Microsoft Word maps the content control to a custom XML node in the document's custom XML data store.
    - `XPath As String` (required): Specifies an XPath string that represents the XML node to which to map the content control. An invalid XPath string causes a run-time error.
    - `PrefixMapping As String` (optional): Specifies the prefix mappings to use when querying the expression provided in the XPath parameter. If omitted, Word uses the set of prefix mappings for the specified custom XML part in the current document.
    - `Source As CustomXMLPart` (optional): Specifies the desired custom XML data to which to map the content control. If this parameter is omitted, the XPath is evaluated against all custom XML in the current document, and the mapping is established with the first CustomXMLPart in which the XPath resolves to an XML node.
- `Delete()`  
  Deletes the XML mapping from the parent content control.
- `SetMappingByNode(Node As CustomXMLNode) As Boolean`  
  Allows creating or changing the XML data mapping on a content control. Returns True if Microsoft Word maps the content control to a custom XML node in the document's custom XML data store.
    - `Node As CustomXMLNode` (required): Specifies the XML node to which to map the current content control.
