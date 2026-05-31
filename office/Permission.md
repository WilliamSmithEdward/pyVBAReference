# Permission

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C0376-0000-0000-C000-000000000046}  

## Properties (16)

- `Application As Object  (read-only)`
- `Creator As Long  (read-only)`
- `Item As UserPermission  (read-only)`
- `Count As Long  (read-only)`
- `EnableTrustedBrowser As Boolean  (read/write)`
- `Parent As Object  (read-only)`
- `Enabled As Boolean  (read/write)`
- `RequestPermissionURL As String  (read/write)`
- `PolicyName As String  (read-only)`
- `PolicyDescription As String  (read-only)`
- `StoreLicenses As Boolean  (read/write)`
- `DocumentAuthor As String  (read/write)`
- `PermissionFromPolicy As Boolean  (read-only)`
- `_NewEnum As IUnknown  (read-only)`
- `DoubleKeyEncryptionUrl As String  (read/write)`
- `SensitivityLabelId As String  (read/write)`

## Methods (3)

- `Add(UserId As String, [Permission As Variant], [ExpirationDate As Variant]) As UserPermission`
- `ApplyPolicy(FileName As String)`
- `RemoveAll()`
