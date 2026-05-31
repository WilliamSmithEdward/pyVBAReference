# Recordset

**Type:** Class  
**Library:** Microsoft ActiveX Data Objects 6.1 Library  
**GUID:** {00000535-0000-0010-8000-00AA006D2EA4}  

## Properties (28)

- `Properties As Properties  (read-only)`
- `AbsolutePosition As PositionEnum  (read/write)`
- `ActiveConnection As Variant  (read/write)`
- `BOF As Boolean  (read-only)`
- `Bookmark As Variant  (read/write)`
- `CacheSize As Long  (read/write)`
- `CursorType As CursorTypeEnum  (read/write)`
- `EOF As Boolean  (read-only)`
- `Fields As Fields  (read-only)`
- `LockType As LockTypeEnum  (read/write)`
- `MaxRecords As Long  (read/write)`
- `RecordCount As Long  (read-only)`
- `Source As Variant  (read/write)`
- `AbsolutePage As PositionEnum  (read/write)`
- `EditMode As EditModeEnum  (read-only)`
- `Filter As Variant  (read/write)`
- `PageCount As Long  (read-only)`
- `PageSize As Long  (read/write)`
- `Sort As String  (read/write)`
- `Status As Long  (read-only)`
- `State As Long  (read-only)`
- `CursorLocation As CursorLocationEnum  (read/write)`
- `MarshalOptions As MarshalOptionsEnum  (read/write)`
- `DataSource As IUnknown  (read/write)`
- `ActiveCommand As Object  (read-only)`
- `StayInSync As Boolean  (read/write)`
- `DataMember As String  (read/write)`
- `Index As String  (read/write)`

## Methods (25)

- `AddNew([FieldList As Variant], [Values As Variant])`
- `CancelUpdate()`
- `Close()`
- `Delete([AffectRecords As AffectEnum])`
- `GetRows([Rows As Long], [Start As Variant], [Fields As Variant]) As Variant`
- `Move(NumRecords As Long, [Start As Variant])`
- `MoveNext()`
- `MovePrevious()`
- `MoveFirst()`
- `MoveLast()`
- `Open([Source As Variant], [ActiveConnection As Variant], [CursorType As CursorTypeEnum], [LockType As LockTypeEnum], [Options As Long])`
- `Requery([Options As Long])`
- `Update([Fields As Variant], [Values As Variant])`
- `UpdateBatch([AffectRecords As AffectEnum])`
- `CancelBatch([AffectRecords As AffectEnum])`
- `NextRecordset([RecordsAffected As Variant]) As _Recordset`
- `Supports(CursorOptions As CursorOptionEnum) As Boolean`
- `Find(Criteria As String, [SkipRecords As Long], [SearchDirection As SearchDirectionEnum], [Start As Variant])`
- `Cancel()`
- `GetString([StringFormat As StringFormatEnum], [NumRows As Long], [ColumnDelimeter As String], [RowDelimeter As String], [NullExpr As String]) As String`
- `CompareBookmarks(Bookmark1 As Variant, Bookmark2 As Variant) As CompareEnum`
- `Clone([LockType As LockTypeEnum]) As _Recordset`
- `Resync([AffectRecords As AffectEnum], [ResyncValues As ResyncEnum])`
- `Seek(KeyValues As Variant, [SeekOption As SeekEnum])`
- `Save([Destination As Variant], [PersistFormat As PersistFormatEnum])`

## Events (11)

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
