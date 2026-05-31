# RecordsetEvents

**Type:** Dispatch Interface  
**Library:** Microsoft ActiveX Data Objects 6.1 Library  
**GUID:** {00001266-0000-0010-8000-00AA006D2EA4}  

## Methods (11)

- `WillChangeField(cFields As Long, Fields As Variant, adStatus As EventStatusEnum, pRecordset As _Recordset)`
- `FieldChangeComplete(cFields As Long, Fields As Variant, pError As Error, adStatus As EventStatusEnum, pRecordset As _Recordset)`
- `WillChangeRecord(adReason As EventReasonEnum, cRecords As Long, adStatus As EventStatusEnum, pRecordset As _Recordset)`
- `RecordChangeComplete(adReason As EventReasonEnum, cRecords As Long, pError As Error, adStatus As EventStatusEnum, pRecordset As _Recordset)`
- `WillChangeRecordset(adReason As EventReasonEnum, adStatus As EventStatusEnum, pRecordset As _Recordset)`
- `RecordsetChangeComplete(adReason As EventReasonEnum, pError As Error, adStatus As EventStatusEnum, pRecordset As _Recordset)`
- `WillMove(adReason As EventReasonEnum, adStatus As EventStatusEnum, pRecordset As _Recordset)`
- `MoveComplete(adReason As EventReasonEnum, pError As Error, adStatus As EventStatusEnum, pRecordset As _Recordset)`
- `EndOfRecordset(fMoreData As Boolean, adStatus As EventStatusEnum, pRecordset As _Recordset)`
- `FetchProgress(Progress As Long, MaxProgress As Long, adStatus As EventStatusEnum, pRecordset As _Recordset)`
- `FetchComplete(pError As Error, adStatus As EventStatusEnum, pRecordset As _Recordset)`
