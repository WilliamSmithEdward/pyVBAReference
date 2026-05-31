# XmlNamespace

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024476-0000-0000-C000-000000000046}  

Represents a namespace that has been added to a workbook.

**Remarks:** Use the Prefix property to return the prefix of an XmlNamespace object. Use the Uri property to return the Uniform Resource Identifier (URI) of an XmlNamespace object.

## Properties (6)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `_Default As String  (read-only)`
- `Uri As String  (read-only)`  
  Returns a String that represents the Uniform Resource Identifier (URI) for the specified XML namespace. Read-only.
- `Prefix As String  (read-only)`  
  Returns a String that represents the prefix for the specified XML namespace. Read-only.
