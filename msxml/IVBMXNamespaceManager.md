# IVBMXNamespaceManager

**Type:** Dispatch Interface  
**Library:** Microsoft XML, v6.0  
**GUID:** {C90352F5-643C-4FBC-BB23-E996EB2D51FD}  

IVBMXNamespaceManager interface

## Properties (1)

- `allowOverride As Boolean  (read/write)`

## Methods (9)

- `reset()`
- `pushContext()`
- `pushNodeContext(contextNode As IXMLDOMNode, [fDeep As Boolean])`
- `popContext()`
- `declarePrefix(prefix As String, namespaceURI As String)`
- `getDeclaredPrefixes() As IMXNamespacePrefixes`
- `getPrefixes(namespaceURI As String) As IMXNamespacePrefixes`
- `getURI(prefix As String) As Variant`
- `getURIFromNode(strPrefix As String, contextNode As IXMLDOMNode) As Variant`
