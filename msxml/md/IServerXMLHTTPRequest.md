# IServerXMLHTTPRequest

**Type:** Dispatch Interface  
**Library:** Microsoft XML, v6.0  
**GUID:** {2E9196BF-13BA-4DD4-91CA-6C571F281495}  

IServerXMLHTTPRequest Interface

## Properties (8)

- `status As Long  (read-only)`  
  Get HTTP status code
- `statusText As String  (read-only)`  
  Get HTTP status text
- `responseXML As Object  (read-only)`  
  Get response body
- `responseText As String  (read-only)`  
  Get response body
- `responseBody As Variant  (read-only)`  
  Get response body
- `responseStream As Variant  (read-only)`  
  Get response body
- `readyState As Long  (read-only)`  
  Get ready state
- `onreadystatechange As Object  (write-only)`  
  Register a complete event handler

## Methods (10)

- `open(bstrMethod As String, bstrUrl As String, [varAsync As Variant], [bstrUser As Variant], [bstrPassword As Variant])`  
  Open HTTP connection
- `setRequestHeader(bstrHeader As String, bstrValue As String)`  
  Add HTTP request header
- `getResponseHeader(bstrHeader As String) As String`  
  Get HTTP response header
- `getAllResponseHeaders() As String`  
  Get all HTTP response headers
- `send([varBody As Variant])`  
  Send HTTP request
- `abort()`  
  Abort HTTP request
- `setTimeouts(resolveTimeout As Long, connectTimeout As Long, sendTimeout As Long, receiveTimeout As Long)`  
  Specify timeout settings (in milliseconds)
- `waitForResponse([timeoutInSeconds As Variant]) As Boolean`  
  Wait for asynchronous send to complete, with optional timeout (in seconds)
- `getOption(option As SERVERXMLHTTP_OPTION) As Variant`  
  Get an option value
- `setOption(option As SERVERXMLHTTP_OPTION, value As Variant)`  
  Set an option value
