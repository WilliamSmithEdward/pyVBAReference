# XSLTransforms

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {C774F5EA-A539-4284-A1BE-30AEC052D899}  

## Properties (5)

- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of XSLTransforms in the collection. Read-only.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified XSLTransforms object.

## Methods (2)

- `Item(Index As Variant) As XSLTransform`  
  Returns an XSLTransform object in a collection.
    - `Index As Variant` (required): The individual object to be returned. Can be a Long indicating the ordinal position or a String representing the name of the individual object.
- `Add(Location As String, [Alias As Variant], [InstallForAllUsers As Boolean]) As XSLTransform`  
  Returns an XSLTransform object that represents an Extensible Stylesheet Language Transformation (XSLT) added to the collection of XSLTs for a specified schema.
    - `Location As String` (required): The path and file name of the XSLT. This may be a local file path, a network path, or an Internet address.
    - `Alias As Variant` (optional): The name of the XSLT as it appears in the Schema Library.
    - `InstallForAllUsers As Boolean` (optional): True if all users that log on to a computer can access and use the new schema. The default is False.
