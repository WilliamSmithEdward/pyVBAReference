# IXMLHTTPRequest

**Type:** Dispatch Interface  
**Library:** Microsoft XML, v6.0  
**GUID:** {ED8C108D-4349-11D2-91A4-00C04F7969E8}  

IXMLHTTPRequest Interface

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

## Methods (6)

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
