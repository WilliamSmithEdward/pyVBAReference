# XmlSchemas

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {0002447A-0000-0000-C000-000000000046}  

Represents the collection of XmlSchema objects contained by an XmlMap object.

**Remarks:** Use the Schemas property of the XmlMap object to return the XmlSchemas collection. Use the Item property to return an XmlSchema object from the XmlSchemas collection.

## Properties (7)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `_Default As XmlSchema  (read-only)`
- `Item As XmlSchema  (read-only)`  
  Returns a single object from a collection.
- `Count As Long  (read-only)`  
  Returns a Long value that represents the number of objects in the collection.
- `_NewEnum As IUnknown  (read-only)`
