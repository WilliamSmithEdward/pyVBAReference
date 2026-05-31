# IVBSAXXMLReader

**Type:** Dispatch Interface  
**Library:** Microsoft XML, v6.0  
**GUID:** {8C033CAA-6CD6-4F73-B728-4531AF74945F}  

IVBSAXXMLReader interface

## Properties (6)

- `entityResolver As IVBSAXEntityResolver  (read/write)`  
  Allow an application to register an entity resolver or look up the current entity resolver.
- `contentHandler As IVBSAXContentHandler  (read/write)`  
  Allow an application to register a content event handler or look up the current content event handler.
- `dtdHandler As IVBSAXDTDHandler  (read/write)`  
  Allow an application to register a DTD event handler or look up the current DTD event handler.
- `errorHandler As IVBSAXErrorHandler  (read/write)`  
  Allow an application to register an error event handler or look up the current error event handler.
- `baseURL As String  (read/write)`  
  Set or get the base URL for the document.
- `secureBaseURL As String  (read/write)`  
  Set or get the secure base URL for the document.

## Methods (6)

- `getFeature(strName As String) As Boolean`  
  Look up the value of a feature.
- `putFeature(strName As String, fValue As Boolean)`  
  Set the state of a feature.
- `getProperty(strName As String) As Variant`  
  Look up the value of a property.
- `putProperty(strName As String, varValue As Variant)`  
  Set the value of a property.
- `parse([varInput As Variant])`  
  Parse an XML document.
- `parseURL(strURL As String)`  
  Parse an XML document from a system identifier (URI).
