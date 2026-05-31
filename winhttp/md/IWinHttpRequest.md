# IWinHttpRequest

**Type:** Dispatch Interface  
**Library:** Microsoft WinHTTP Services, version 5.1  
**GUID:** {016FE2EC-B2C8-45F8-B23B-39E53A75396B}  

IWinHttpRequest Interface

## Properties (6)

- `Status As Long  (read-only)`  
  Get HTTP status code
- `StatusText As String  (read-only)`  
  Get HTTP status text
- `ResponseText As String  (read-only)`  
  Get response body as a string
- `ResponseBody As Variant  (read-only)`  
  Get response body as a safearray of UI1
- `ResponseStream As Variant  (read-only)`  
  Get response body as a stream
- `Option As Variant  (read/write)`

## Methods (12)

- `SetProxy(ProxySetting As HTTPREQUEST_PROXY_SETTING, [ProxyServer As Variant], [BypassList As Variant])`  
  Specify proxy configuration
- `SetCredentials(UserName As String, Password As String, Flags As HTTPREQUEST_SETCREDENTIALS_FLAGS)`  
  Specify authentication credentials
- `Open(Method As String, Url As String, [Async As Variant])`  
  Open HTTP connection
- `SetRequestHeader(Header As String, Value As String)`  
  Add HTTP request header
- `GetResponseHeader(Header As String) As String`  
  Get HTTP response header
- `GetAllResponseHeaders() As String`  
  Get all HTTP response headers
- `Send([Body As Variant])`  
  Send HTTP request
- `WaitForResponse([Timeout As Variant]) As Boolean`  
  Wait for asynchronous send to complete, with optional timeout (in seconds)
- `Abort()`  
  Abort an asynchronous operation in progress
- `SetTimeouts(ResolveTimeout As Long, ConnectTimeout As Long, SendTimeout As Long, ReceiveTimeout As Long)`  
  Specify timeout settings (in milliseconds)
- `SetClientCertificate(ClientCertificate As String)`  
  Specify a client certificate
- `SetAutoLogonPolicy(AutoLogonPolicy As WinHttpRequestAutoLogonPolicy)`  
  Specify if credentials should be sent automatically
