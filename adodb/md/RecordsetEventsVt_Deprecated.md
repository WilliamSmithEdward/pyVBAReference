# RecordsetEventsVt_Deprecated

**Type:** Interface  
**Library:** Microsoft ActiveX Data Objects 6.1 Library  
**GUID:** {00000403-0000-0010-8000-00AA006D2EA4}  

## Methods (11)

- `WillChangeField(cFields As Long, Fields As Variant, adStatus As EventStatusEnum, pRecordset As _Recordset_Deprecated)`
- `FieldChangeComplete(cFields As Long, Fields As Variant, pError As Error, adStatus As EventStatusEnum, pRecordset As _Recordset_Deprecated)`
- `WillChangeRecord(adReason As EventReasonEnum, cRecords As Long, adStatus As EventStatusEnum, pRecordset As _Recordset_Deprecated)`
- `RecordChangeComplete(adReason As EventReasonEnum, cRecords As Long, pError As Error, adStatus As EventStatusEnum, pRecordset As _Recordset_Deprecated)`
- `WillChangeRecordset(adReason As EventReasonEnum, adStatus As EventStatusEnum, pRecordset As _Recordset_Deprecated)`
- `RecordsetChangeComplete(adReason As EventReasonEnum, pError As Error, adStatus As EventStatusEnum, pRecordset As _Recordset_Deprecated)`
- `WillMove(adReason As EventReasonEnum, adStatus As EventStatusEnum, pRecordset As _Recordset_Deprecated)`
- `MoveComplete(adReason As EventReasonEnum, pError As Error, adStatus As EventStatusEnum, pRecordset As _Recordset_Deprecated)`
- `EndOfRecordset(fMoreData As Boolean, adStatus As EventStatusEnum, pRecordset As _Recordset_Deprecated)`
- `FetchProgress(Progress As Long, MaxProgress As Long, adStatus As EventStatusEnum, pRecordset As _Recordset_Deprecated)`
- `FetchComplete(pError As Error, adStatus As EventStatusEnum, pRecordset As _Recordset_Deprecated)`
