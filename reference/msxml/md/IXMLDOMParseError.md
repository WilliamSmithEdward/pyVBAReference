# IXMLDOMParseError

**Type:** Dispatch Interface  
**Library:** Microsoft XML, v6.0  
**GUID:** {3EFAA426-272F-11D2-836F-0000F87A7782}  

structure for reporting parser errors

## Properties (7)

- `errorCode As Long  (read-only)`  
  the error code
- `url As String  (read-only)`  
  the URL of the XML document containing the error
- `reason As String  (read-only)`  
  the cause of the error
- `srcText As String  (read-only)`  
  the data where the error occurred
- `line As Long  (read-only)`  
  the line number in the XML document where the error occurred
- `linepos As Long  (read-only)`  
  the character position in the line containing the error
- `filepos As Long  (read-only)`  
  the absolute file position in the XML document containing the error
