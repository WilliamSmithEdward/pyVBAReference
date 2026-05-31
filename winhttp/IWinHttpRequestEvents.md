# IWinHttpRequestEvents

**Type:** Interface  
**Library:** Microsoft WinHTTP Services, version 5.1  
**GUID:** {F97F4E15-B787-4212-80D1-D380CBBF982E}  

IWinHttpRequestEvents Interface

## Methods (4)

- `OnResponseStart(Status As Long, ContentType As String)`
- `OnResponseDataAvailable(Data As SAFEARRAY(Byte))`
- `OnResponseFinished()`
- `OnError(ErrorNumber As Long, ErrorDescription As String)`
