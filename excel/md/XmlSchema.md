# XmlSchema

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024479-0000-0000-C000-000000000046}  

Represents an XML schema contained by an XmlMap object.

**Remarks:** Use the Item method to return an XmlSchema object from the XmlSchemas collection. Use the Namespace property to return the target namespace for a schema. Use the XML property to return the XML contents of a schema.

## Properties (6)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Namespace As XmlNamespace  (read-only)`  
  Returns an XmlNamespace object that represents the target namespace for the specified schema. Read-only.
- `XML As String  (read-only)`  
  Returns a String value that represents the content of the specified schema.
- `Name As String  (read-only)`  
  Returns a String value that represents the friendly name used to identify an XML schema in an XmlMap object.
