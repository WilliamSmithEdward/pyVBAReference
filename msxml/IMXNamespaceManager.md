# IMXNamespaceManager

**Type:** Interface  
**Library:** Microsoft XML, v6.0  
**GUID:** {C90352F6-643C-4FBC-BB23-E996EB2D51FD}  

IMXNamespaceManager interface

## Methods (10)

- `putAllowOverride(fOverride As Boolean)`
- `getAllowOverride(fOverride As Boolean)`
- `reset()`
- `pushContext()`
- `pushNodeContext(contextNode As IXMLDOMNode, fDeep As Boolean)`
- `popContext()`
- `declarePrefix(prefix As String, namespaceURI As String)`
- `getDeclaredPrefix(nIndex As Long, pwchPrefix As Integer, pcchPrefix As Long)`
- `getPrefix(pwszNamespaceURI As String, nIndex As Long, pwchPrefix As Integer, pcchPrefix As Long)`
- `getURI(pwchPrefix As String, pContextNode As IXMLDOMNode, pwchUri As Integer, pcchUri As Long)`
