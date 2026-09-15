# RecordsetEvents

**Type:** Dispatch Interface  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {45165490-EF32-11D0-86FB-006097C9818C}  

## Methods (11)

- `WillChangeField(cFields As Long, Fields As Variant, adStatus As EventStatusEnum, pRecordset As IUnknown)`
- `FieldChangeComplete(cFields As Long, Fields As Variant, pError As IUnknown, adStatus As EventStatusEnum, pRecordset As IUnknown)`
- `WillChangeRecord(adReason As EventReasonEnum, cRecords As Long, adStatus As EventStatusEnum, pRecordset As IUnknown)`
- `RecordChangeComplete(adReason As EventReasonEnum, cRecords As Long, pError As IUnknown, adStatus As EventStatusEnum, pRecordset As IUnknown)`
- `WillChangeRecordset(adReason As EventReasonEnum, adStatus As EventStatusEnum, pRecordset As IUnknown)`
- `RecordsetChangeComplete(adReason As EventReasonEnum, pError As IUnknown, adStatus As EventStatusEnum, pRecordset As IUnknown)`
- `WillMove(adReason As EventReasonEnum, adStatus As EventStatusEnum, pRecordset As IUnknown)`
- `MoveComplete(adReason As EventReasonEnum, pError As IUnknown, adStatus As EventStatusEnum, pRecordset As IUnknown)`
- `EndOfRecordset(fMoreData As Integer, adStatus As EventStatusEnum, pRecordset As IUnknown)`
- `FetchProgress(Progress As Long, MaxProgress As Long, adStatus As EventStatusEnum, pRecordset As IUnknown)`
- `FetchComplete(pError As IUnknown, adStatus As EventStatusEnum, pRecordset As IUnknown)`
