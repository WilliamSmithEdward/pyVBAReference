# IVBSAXErrorHandler

**Type:** Dispatch Interface  
**Library:** Microsoft XML, v6.0  
**GUID:** {D963D3FE-173C-4862-9095-B92F66995F52}  

IVBSAXErrorHandler interface

## Methods (3)

- `error(oLocator As IVBSAXLocator, strErrorMessage As String, nErrorCode As Long)`  
  Receive notification of a recoverable error.
- `fatalError(oLocator As IVBSAXLocator, strErrorMessage As String, nErrorCode As Long)`  
  Receive notification of a non-recoverable error.
- `ignorableWarning(oLocator As IVBSAXLocator, strErrorMessage As String, nErrorCode As Long)`  
  Receive notification of an ignorable warning.
