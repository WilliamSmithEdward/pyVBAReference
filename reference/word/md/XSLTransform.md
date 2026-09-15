# XSLTransform

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {E3124493-7D6A-410F-9A48-CC822C033CEC}  

Represents a single registered Extensible Stylesheet Language Transformation (XSLT).

**Remarks:** Use the Add method of the XSLTransforms collection to add an individual XSLT to the list of XSLTs available for a schema. The following example adds the simplesample.xslt transformation to the XSLTs for the SimpleSample schema. Use the Item method of the XSLTransforms collection to return a single XSLTransform object. The following example deletes the first XSLT in the collection of XSLTs for the SimpleSample schema.

## Properties (6)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified XSLTransform object.
- `Alias As String  (read/write)`  
  Returns a String that represents the display name for the specified object.
- `Location As String  (read/write)`  
  Returns or sets a String that represents the physical location of the XSL transform for the namespace of an XML schema in the Schema Library. Read/write.
- `ID As String  (read-only)`  
  Returns a String containing the GUID assigned to the current XSLTransform object. Read-only.

## Methods (1)

- `Delete()`  
  Deletes the specified Extensible Stylesheet Language Transformation (XSLT) from the list of avaliable XSLTs.
