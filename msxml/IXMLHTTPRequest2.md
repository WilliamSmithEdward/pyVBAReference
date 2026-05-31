# IXMLHTTPRequest2

**Type:** Interface  
**Library:** Microsoft XML, v6.0  
**GUID:** {E5D37DC0-552A-4D52-9CC0-A14D546FBD04}  

IXMLHTTPRequest2 Interface

## Methods (10)

- `open(pwszMethod As String, pwszUrl As String, pStatusCallback As IXMLHTTPRequest2Callback, pwszUserName As String, pwszPassword As String, pwszProxyUserName As String, pwszProxyPassword As String)`
- `send(pBody As ISequentialStream, cbBody As LongLong)`
- `abort()`
- `SetCookie(pCookie As tagXHR_COOKIE, pdwCookieState As Long)`
- `SetCustomResponseStream(pSequentialStream As ISequentialStream)`
- `setProperty(eProperty As XHR_PROPERTY, ullValue As LongLong)`
- `setRequestHeader(pwszHeader As String, pwszValue As String)`
- `getAllResponseHeaders(ppwszHeaders As String)`
- `GetCookie(pwszUrl As String, pwszName As String, dwFlags As Long, pcCookies As Long, ppCookies As tagXHR_COOKIE)`
- `getResponseHeader(pwszHeader As String, ppwszValue As String)`
