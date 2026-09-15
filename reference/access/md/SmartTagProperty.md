# SmartTagProperty

**Type:** Class  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {6E03AD86-431E-4879-A572-EF0EBA2FA729}  

Represents a single instance of a custom property for a smart tag. The SmartTagProperty object is a member of the SmartTagProperties collection.

**Remarks:** To return a SmartTagProperty object, use the Item method of the SmartTagProperties collection, or use SmartTagProperties (index), where index is the number of the property. To return the information related to a custom property for a smart tag, use the Name and Value properties of the SmartTagProperty object.

## Properties (2)

- `Name As String  (read/write)`  
  Use the Name property to specify or determine the string expression that identifies the name of an object. Read/write String.
- `Value As String  (read/write)`  
  Gets or sets the value of the specified object. Read/write Variant.

## Methods (1)

- `Delete()`  
  Deletes the specified object.
