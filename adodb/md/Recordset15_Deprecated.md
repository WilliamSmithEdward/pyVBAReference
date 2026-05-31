# Recordset15_Deprecated

**Type:** Dispatch Interface  
**Library:** Microsoft ActiveX Data Objects 6.1 Library  
**GUID:** {0000050E-0000-0010-8000-00AA006D2EA4}  

## Properties (23)

- `Properties As Properties  (read-only)`
- `AbsolutePosition As PositionEnum_Param  (read/write)`
- `ActiveConnection As Variant  (read/write)`
- `BOF As Boolean  (read-only)`
- `Bookmark As Variant  (read/write)`
- `CacheSize As Long  (read/write)`
- `CursorType As CursorTypeEnum  (read/write)`
- `EOF As Boolean  (read-only)`
- `Fields As Fields_Deprecated  (read-only)`
- `LockType As LockTypeEnum  (read/write)`
- `MaxRecords As ADO_LONGPTR  (read/write)`
- `RecordCount As ADO_LONGPTR  (read-only)`
- `Source As Variant  (read/write)`
- `AbsolutePage As PositionEnum_Param  (read/write)`
- `EditMode As EditModeEnum  (read-only)`
- `Filter As Variant  (read/write)`
- `PageCount As ADO_LONGPTR  (read-only)`
- `PageSize As Long  (read/write)`
- `Sort As String  (read/write)`
- `Status As Long  (read-only)`
- `State As Long  (read-only)`
- `CursorLocation As CursorLocationEnum  (read/write)`
- `MarshalOptions As MarshalOptionsEnum  (read/write)`

## Methods (18)

- `AddNew([FieldList As Variant], [Values As Variant])`
- `CancelUpdate()`
- `Close()`
- `Delete([AffectRecords As AffectEnum])`
- `GetRows([Rows As Long], [Start As Variant], [Fields As Variant]) As Variant`
- `Move(NumRecords As ADO_LONGPTR, [Start As Variant])`
- `MoveNext()`
- `MovePrevious()`
- `MoveFirst()`
- `MoveLast()`
- `Open([Source As Variant], [ActiveConnection As Variant], [CursorType As CursorTypeEnum], [LockType As LockTypeEnum], [Options As Long])`
- `Requery([Options As Long])`
- `Update([Fields As Variant], [Values As Variant])`
- `UpdateBatch([AffectRecords As AffectEnum])`
- `CancelBatch([AffectRecords As AffectEnum])`
- `NextRecordset([RecordsAffected As Variant]) As _Recordset_Deprecated`
- `Supports(CursorOptions As CursorOptionEnum) As Boolean`
- `Find(Criteria As String, [SkipRecords As ADO_LONGPTR], [SearchDirection As SearchDirectionEnum], [Start As Variant])`
