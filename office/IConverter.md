# IConverter

**Type:** Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C03D7-0000-0000-C000-000000000046}  

## Methods (6)

- `HrInitConverter(pcap As IConverterApplicationPreferences, ppcp As IConverterPreferences, pcuic As IConverterUICallback)`
- `HrUninitConverter(pcuic As IConverterUICallback)`
- `HrImport(bstrSourcePath As String, bstrDestPath As String, pcap As IConverterApplicationPreferences, ppcp As IConverterPreferences, pcuic As IConverterUICallback)`
- `HrExport(bstrSourcePath As String, bstrDestPath As String, bstrClass As String, pcap As IConverterApplicationPreferences, ppcp As IConverterPreferences, pcuic As IConverterUICallback)`
- `HrGetFormat(bstrPath As String, pbstrClass As String, pcap As IConverterApplicationPreferences, ppcp As IConverterPreferences, pcuic As IConverterUICallback)`
- `HrGetErrorString(hrErr As Long, pbstrErrorMsg As String, pcap As IConverterApplicationPreferences)`
