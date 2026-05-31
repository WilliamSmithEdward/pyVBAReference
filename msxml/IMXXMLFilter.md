# IMXXMLFilter

**Type:** Dispatch Interface  
**Library:** Microsoft XML, v6.0  
**GUID:** {C90352F7-643C-4FBC-BB23-E996EB2D51FD}  

IMXXMLFilter interface

## Properties (4)

- `entityResolver As IUnknown  (read/write)`
- `contentHandler As IUnknown  (read/write)`
- `dtdHandler As IUnknown  (read/write)`
- `errorHandler As IUnknown  (read/write)`

## Methods (4)

- `getFeature(strName As String) As Boolean`
- `putFeature(strName As String, fValue As Boolean)`
- `getProperty(strName As String) As Variant`
- `putProperty(strName As String, varValue As Variant)`
