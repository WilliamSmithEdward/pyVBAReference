# SmartTagProperties

**Type:** Class  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {4215CC2C-15B5-47A5-9B60-119BD269CB7E}  

A collection of SmartTagProperty objects that represents the properties related to a smart tag.

**Remarks:** To return the SmartTagProperties collection for a smart tag, use the Properties property of the SmartTag object. To create a custom property from within a Visual Basic for Applications (VBA) project, use the Add method of the SmartTagProperties collection.

## Properties (4)

- `Application As Application  (read-only)`  
  Use the Application property to access the active Microsoft Access Application object and its related properties. Read-only Application object.
- `Count As Long  (read-only)`  
  Use the Count property to determine the number of items in a specified collection. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Item As _SmartTagProperty  (read-only)`  
  The Item property returns a specific member of a collection either by position or by index. Read-only SmartTagProperty.

## Methods (1)

- `Add(Name As String, Value As Variant) As _SmartTagProperty`  
  Adds a custom property to a smart tag.
    - `Name As String` (required): The name to be used for the custom property.
    - `Value As Variant` (required): The value of the custom property.
