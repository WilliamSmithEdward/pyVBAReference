# ILicWizExternal

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {4CAC6328-B9B0-11D3-8D59-0050048384E3}  

## Properties (9)

- `Context As Long  (read-only)`
- `Validator As Object  (read-only)`
- `LicAgent As Object  (read-only)`
- `CountryInfo As String  (read-only)`
- `WizardVisible As Long  (write-only)`
- `WizardTitle As String  (write-only)`
- `AnimationEnabled As Long  (read-only)`
- `CurrentHelpId As Long  (write-only)`
- `OfficeOnTheWebUrl As String  (read-only)`

## Methods (18)

- `PrintHtmlDocument(punkHtmlDoc As IUnknown)`
- `InvokeDateTimeApplet()`
- `FormatDate(date As Date, [pFormat As String]) As String`
- `ShowHelp([pvarId As Variant])`
- `Terminate()`
- `DisableVORWReminder(BPC As Long)`
- `SaveReceipt(bstrReceipt As String) As String`
- `OpenInDefaultBrowser(bstrUrl As String)`
- `MsoAlert(bstrText As String, bstrButtons As String, bstrIcon As String) As Long`
- `DepositPidKey(bstrKey As String, fMORW As Long) As Long`
- `WriteLog(bstrMessage As String)`
- `ResignDpc(bstrProductCode As String)`
- `ResetPID()`
- `SetDialogSize(dx As Long, dy As Long)`
- `VerifyClock(lMode As Long) As Long`
- `SortSelectOptions(pdispSelect As Object)`
- `InternetDisconnect()`
- `GetConnectedState() As Long`
