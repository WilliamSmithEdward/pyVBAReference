# FreeThreadedXMLHTTP60

**Type:** Class  
**Library:** Microsoft XML, v6.0  
**GUID:** {88D96A09-F192-11D4-A65F-0040963251E5}  

Free Threaded XML HTTP Request class 6.0

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
