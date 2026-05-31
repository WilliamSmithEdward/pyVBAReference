# ISAXXMLReader

**Type:** Interface  
**Library:** Microsoft XML, v6.0  
**GUID:** {A4F96ED0-F829-476E-81C0-CDC7BD2A0802}  

ISAXXMLReader interface

## Methods (18)

- `getFeature(pwchName As Integer, pvfValue As Boolean)`
- `putFeature(pwchName As Integer, vfValue As Boolean)`
- `getProperty(pwchName As Integer, pvarValue As Variant)`
- `putProperty(pwchName As Integer, varValue As Variant)`
- `getEntityResolver(ppResolver As ISAXEntityResolver)`
- `putEntityResolver(pResolver As ISAXEntityResolver)`
- `getContentHandler(ppHandler As ISAXContentHandler)`
- `putContentHandler(pHandler As ISAXContentHandler)`
- `getDTDHandler(ppHandler As ISAXDTDHandler)`
- `putDTDHandler(pHandler As ISAXDTDHandler)`
- `getErrorHandler(ppHandler As ISAXErrorHandler)`
- `putErrorHandler(pHandler As ISAXErrorHandler)`
- `getBaseURL(ppwchBaseUrl As Integer)`
- `putBaseURL(pwchBaseUrl As Integer)`
- `getSecureBaseURL(ppwchSecureBaseUrl As Integer)`
- `putSecureBaseURL(pwchSecureBaseUrl As Integer)`
- `parse([varInput As Variant])`
- `parseURL(pwchUrl As Integer)`
