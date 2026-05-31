# XmlNamespaces

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024477-0000-0000-C000-000000000046}  

Represents the collection of XmlNamespace objects in a workbook.

**Remarks:** Use the Item method to access a particular XmlNamespace object. Use the Value property to return a String that lists the namespaces that have been added to a workbook.

## Properties (8)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `_Default As XmlNamespace  (read-only)`
- `Item As XmlNamespace  (read-only)`  
  Returns a single object from a collection.
- `Count As Long  (read-only)`  
  Returns a Long value that represents the number of objects in the collection.
- `Value As String  (read-only)`  
  Returns a String value that represents the XML namespaces that have been added to the workbook.
- `_NewEnum As IUnknown  (read-only)`

## Methods (1)

- `InstallManifest(Path As String, [InstallForAllUsers As Variant])`  
  Installs the specified XML expansion pack on the user's computer, making an XML smart document solution available to one or more users.
    - `Path As String` (required): The path and file name of the XML expansion pack.
    - `InstallForAllUsers As Variant` (optional): True installs the XML expansion pack and makes it available to all users on a machine. False makes the XML expansion pack available for the current user only. The default is False.
