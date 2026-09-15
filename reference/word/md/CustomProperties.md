# CustomProperties

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {B923FDE1-F08C-11D3-91B0-00105A0A19FD}  

A collection of CustomProperty objects that represents the properties related to a smart tag. The CustomProperties collection includes all the smart tag custom properties in a document.

**Remarks:** Use the Properties property to return a single CustomProperties object. Use the Add method of the CustomProperties object with to create a custom property from within a Microsoft Word Visual Basic for Applications project. This example creates a new property for the first smart tag in the active document and displays the XML code used for the tag. Use Properties (Index) to return a single property for a smart tag, where Index is the number of the property. This example displays the name and value of the first property of the first smart tag in the current document. Use the Count property to return the number of custom properties for a smart tag. This example loops through all the smart tags in the current document and then lists in a new document the name and value of the custom properties for all smart tags that have custom properties.

## Properties (5)

- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of items in the collection. Read-only.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified CustomProperties object.

## Methods (2)

- `Item(Index As Variant) As CustomProperty`  
  Returns a CustomProperty object in a collection.
    - `Index As Variant` (required): The individual object to be returned. Can be a Long indicating the ordinal position or a String representing the name of the individual object.
- `Add(Name As String, Value As String) As CustomProperty`  
  Returns a CustomProperty object that represents s custom property added to a smart tag.
    - `Name As String` (required): The name of the custom smart tag property.
    - `Value As String` (required): The value of the custom smart tag property
