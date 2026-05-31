# Recordset20

**Type:** Dispatch Interface  
**Library:** Microsoft ActiveX Data Objects 6.1 Library  
**GUID:** {0000154F-0000-0010-8000-00AA006D2EA4}  

## Properties (27)

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

## Methods (23)

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
