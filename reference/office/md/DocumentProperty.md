# DocumentProperty

**Type:** Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {2DF8D04E-5BFA-101B-BDE5-00AA0044DE52}  

Represents a custom or built-in document property of a container document. The DocumentProperty object is a member of the DocumentProperties collection.

**Remarks:** Use the Microsoft Word Document.BuiltinDocumentProperties(_index_) property, where _index_ is the name or index number of the built-in document property, to return a single DocumentProperty object that represents a specific built-in document property. Use the Word Document.CustomDocumentProperties(_index_) property, where _index_ is the name or index number of the custom document property, to return a DocumentProperty object that represents a specific custom document property.

## Properties (8)

- `Parent As Object  (read-only)`  
  Gets the Parent object for the DocumentProperty object. Read-only.
- `Name As HRESULT  (read/write)`  
  Gets or sets the name of a document property. Read/write.
- `Value As HRESULT  (read/write)`  
  Gets or sets the value of a document property. Read/write.
- `Type As HRESULT  (read/write)`  
  Gets or sets the document property type. Read-only for built-in document properties; read/write for custom document properties.
- `LinkToContent As HRESULT  (read/write)`  
  Is True if the value of the custom document property is linked to the content of the container document. False if the value is static. Read/write.
- `LinkSource As HRESULT  (read/write)`  
  Gets or sets the source of a linked custom document property. Read/write.
- `Application As HRESULT  (read-only)`  
  Gets an Application object that represents the container application for the DocumentProperty object (you can use this property with an Automation object to return that object's container application). Read-only.
- `Creator As HRESULT  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the DocumentProperty object was created. Read-only.

## Methods (1)

- `Delete()`  
  Removes a custom document property.
