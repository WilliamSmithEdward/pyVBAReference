# Workbooks

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000208DB-0000-0000-C000-000000000046}  

## Properties (7)

- `Application As Application  (read-only)`
- `Creator As XlCreator  (read-only)`
- `Parent As Object  (read-only)`
- `Count As Long  (read-only)`
- `Item As Workbook  (read-only)`
- `_NewEnum As IUnknown  (read-only)`
- `_Default As Workbook  (read-only)`

## Methods (8)

- `Add([Template As Variant]) As Workbook`
- `Close()`
- `Open(Filename As String, [UpdateLinks As Variant], [ReadOnly As Variant], [Format As Variant], [Password As Variant], [WriteResPassword As Variant], [IgnoreReadOnlyRecommended As Variant], [Origin As Variant], [Delimiter As Variant], [Editable As Variant], [Notify As Variant], [Converter As Variant], [AddToMru As Variant], [Local As Variant], [CorruptLoad As Variant]) As Workbook`
- `OpenText(Filename As String, [Origin As Variant], [StartRow As Variant], [DataType As Variant], [TextQualifier As XlTextQualifier], [ConsecutiveDelimiter As Variant], [Tab As Variant], [Semicolon As Variant], [Comma As Variant], [Space As Variant], [Other As Variant], [OtherChar As Variant], [FieldInfo As Variant], [TextVisualLayout As Variant], [DecimalSeparator As Variant], [ThousandsSeparator As Variant], [TrailingMinusNumbers As Variant], [Local As Variant])`
- `OpenDatabase(Filename As String, [CommandText As Variant], [CommandType As Variant], [BackgroundQuery As Variant], [ImportDataAs As Variant]) As Workbook`
- `CheckOut(Filename As String)`
- `CanCheckOut(Filename As String) As Boolean`
- `OpenXML(Filename As String, [Stylesheets As Variant], [LoadOption As Variant]) As Workbook`
