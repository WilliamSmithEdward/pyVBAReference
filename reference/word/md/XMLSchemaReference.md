# XMLSchemaReference

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {FE0971F0-5E60-4985-BCDA-95CB0B8E0308}  

Represents an individual schema that is attached to a document.

**Remarks:** Use the XMLSchemaReference property to return an XMLSchemaReference object for a ChildNodeSuggestion object. The following example inserts the suggested XML child element if the XML schema referenced is the SimpleSample schema.

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified XMLSchemaReference object.
- `NamespaceURI As String  (read-only)`  
  Returns a String that represents the Uniform Resource Identifier (URI) of the schema namespace for the specified object. Read-only.
- `Location As String  (read-only)`  
  Returns a String that represents the physical location of an XML schema. Read-only.

## Methods (2)

- `Delete()`  
  Deletes the specified XML Schema Reference.
- `Reload()`  
  Reloads the XML schemas that are referenced in a document.
