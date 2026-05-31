# MXNamespaceManager60

**Type:** Class  
**Library:** Microsoft XML, v6.0  
**GUID:** {88D96A11-F192-11D4-A65F-0040963251E5}  

MX Namespace Manager 6.0

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
