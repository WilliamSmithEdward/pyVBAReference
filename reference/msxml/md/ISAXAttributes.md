# ISAXAttributes

**Type:** Interface  
**Library:** Microsoft XML, v6.0  
**GUID:** {F078ABE1-45D2-4832-91EA-4466CE2F25C9}  

ISAXAttributes interface

## Methods (13)

- `getLength(pnLength As Long)`
- `getURI(nIndex As Long, ppwchUri As Integer, pcchUri As Long)`
- `getLocalName(nIndex As Long, ppwchLocalName As Integer, pcchLocalName As Long)`
- `getQName(nIndex As Long, ppwchQName As Integer, pcchQName As Long)`
- `getName(nIndex As Long, ppwchUri As Integer, pcchUri As Long, ppwchLocalName As Integer, pcchLocalName As Long, ppwchQName As Integer, pcchQName As Long)`
- `getIndexFromName(pwchUri As Integer, cchUri As Long, pwchLocalName As Integer, cchLocalName As Long, pnIndex As Long)`
- `getIndexFromQName(pwchQName As Integer, cchQName As Long, pnIndex As Long)`
- `getType(nIndex As Long, ppwchType As Integer, pcchType As Long)`
- `getTypeFromName(pwchUri As Integer, cchUri As Long, pwchLocalName As Integer, cchLocalName As Long, ppwchType As Integer, pcchType As Long)`
- `getTypeFromQName(pwchQName As Integer, cchQName As Long, ppwchType As Integer, pcchType As Long)`
- `getValue(nIndex As Long, ppwchValue As Integer, pcchValue As Long)`
- `getValueFromName(pwchUri As Integer, cchUri As Long, pwchLocalName As Integer, cchLocalName As Long, ppwchValue As Integer, pcchValue As Long)`
- `getValueFromQName(pwchQName As Integer, cchQName As Long, ppwchValue As Integer, pcchValue As Long)`
