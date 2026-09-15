# XMLNamespaces

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {656BBED7-E82D-4B0A-8F97-EC742BA11FFA}  

## Properties (5)

- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of XML namespaces in the collection. Read-only.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified XMLNamespaces object.

## Methods (3)

- `Item(Index As Variant) As XMLNamespace`  
  Returns an individual XMLNamespace object in a collection.
    - `Index As Variant` (required): The individual object to be returned. Can be a Long indicating the ordinal position or a String representing the name of the individual object.
- `Add(Path As String, [NamespaceURI As Variant], [Alias As Variant], [InstallForAllUsers As Boolean]) As XMLNamespace`  
  Returns an XMLNamespace object that represents a schema that is added to the Schema Library and made available to users in Microsoft Word.
    - `Path As String` (required): The path and file name of the schema. This may be a local file path, a network path, or an Internet address.
    - `NamespaceURI As Variant` (optional): The namespace Uniform Resource Indicator as specified in the schema. The NamespaceURI parameter is case-sensitive and must be spelled exactly as specified in schema.
    - `Alias As Variant` (optional): The name of the schema as it appears on the Schemas tab in the Templates and Add-ins dialog box.
    - `InstallForAllUsers As Boolean` (optional): True if all users that log on to a computer can access and use the new schema. The default is False.
- `InstallManifest(Path As String, [InstallForAllUsers As Boolean])`  
  Installs the specified XML expansion pack on the user's computer, making an XML smart document solution available to one or more users.
    - `Path As String` (required): The path and file name of the XML expansion pack.
    - `InstallForAllUsers As Boolean` (optional): True installs the XML expansion pack and makes it available to all users on a computer. False makes the XML expansion pack available for the current user only. Default is False.
