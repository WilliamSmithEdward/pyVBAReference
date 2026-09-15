# XMLNodes

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {D36C1F42-7044-4B9E-9CA3-85919454DB04}  

A collection of XMLNode objects that represents the nodes in the tree view of the XML Structure task pane, which indicates the elements that a user has applied to a document. Each node in the tree view is an instance of an XMLNode object. The hierarchy in the tree view indicates whether a node contains child nodes.Y

**Remarks:** ou can return an XMLNodes collection for a selection, a range, or the entire document. The order in which the XMLNode objects appear in the XMLNodes collection is the same order in which their start or end tags appear within the specified selection, range, or document. Use the Item method of the XMLNodes collection to return an individual XMLNode object. Use the Validate method to verify that an XML element is valid according to the applied schemas and that any required child elements exist and are in the required order. Once you run the Validate method, use the ValidationStatus property to verify whether an element is valid and the ValidationErrorText property to display a message to the user as to what the user needs to fix to make the XML in the document conform to the XML schema rules. The following example validates each of the XML elements in the active document, and if the element or attribute is found to be invalid against the schema, returns a message to the user explaining why the element is invalid. Use the Add method to add an XML element to a selection, a range, or the document.

## Properties (5)

- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of XML elements in the collection. Read-only.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified XMLNodes object.

## Methods (1)

- `Item(Index As Long) As XMLNode`  
  Returns an individual XMLNode object in a collection.
    - `Index As Long` (required): The individual object to be returned. Can be a Long indicating the ordinal position of the individual object.
