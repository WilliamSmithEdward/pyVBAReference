# XMLNamespace

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {B140A023-4850-4DA6-BC5F-CC459C4507BC}  

Represents an individual schema within the Schema Library.

**Remarks:** You can access the Schema Library from within Microsoft Word from the XML Schema tab in the Templates and Add-ins dialog box. The Schema Library represents schemas installed on a user's computer that a user has applied to a Word document or that a user has explicitly added to the Schema Library by using the Schema Library dialog box. Use the Item method of the XMLNamespaces collection to return an individual XMLNameSpace object. The index value of the Item method can be either a Long, which indicates the position of the schema in the Schema Library, or a String, which represents the name of the schema as returned using the URI property (the TargetNamespace setting defined in the schema). The following example attaches a schema named SimpleSample to the active document.

## Properties (8)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified XMLNamespace object.
- `URI As String  (read-only)`
- `Location As String  (read/write)`  
  Returns or sets a String that represents the physical location of the namespace of an XML schema in the Schema Library. Read/write.
- `Alias As String  (read/write)`  
  Returns a String that represents the display name for the specified object.
- `XSLTransforms As XSLTransforms  (read-only)`  
  Returns an XSLTransforms collection that represents the Extensible Stylesheet Language Transformation (XSLT) files specified for use with a schema.
- `DefaultTransform As XSLTransform  (read/write)`  
  Returns an XSLTransform object that represents the default Extensible Stylesheet Language Transformation (XSLT) file to use when opening a document from an XML schema for a particular namespace.

## Methods (2)

- `AttachToDocument(Document As Variant)`  
  Attaches an XML schema to a document.
    - `Document As Variant` (required): The document to which to attach the specified XML schema.
- `Delete()`  
  Deletes the specified XML schema from the list of available XML schemas.
