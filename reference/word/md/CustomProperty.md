# CustomProperty

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {B923FDE0-F08C-11D3-91B0-00105A0A19FD}  

Represents a single instance of a custom property for a smart tag. The CustomProperty object is a member of the CustomProperties collection.

**Remarks:** Use the Item methodor Properties (Index), where Index is the number of the propertyof the CustomProperties collection to return a CustomProperty object. Use the Name and Value properties to return the information related to a custom property for a smart tag. This example displays a message containing the name and value of the first custom property of the first smart tag in the current document. This example assumes that the current document contains at least one smart tag and that the first smart tag has at least one custom property.

## Properties (5)

- `Name As String  (read-only)`  
  Returns the name of the specified object. Read-only String.
- `Value As String  (read/write)`  
  Returns or sets the value of a custom property. Read/write String.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified CustomProperty object.

## Methods (1)

- `Delete()`  
  Deletes the specified custom property.
