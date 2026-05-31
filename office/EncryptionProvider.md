# EncryptionProvider

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000CD809-0000-0000-C000-000000000046}  

## Methods (9)

- `GetProviderDetail(encprovdet As EncryptionProviderDetail) As Variant`
- `NewSession(ParentWindow As IUnknown) As Long`
- `Authenticate(ParentWindow As IUnknown, EncryptionData As IUnknown, PermissionsMask As Long) As Long`
- `CloneSession(SessionHandle As Long) As Long`
- `EndSession(SessionHandle As Long)`
- `Save(SessionHandle As Long, EncryptionData As IUnknown) As Long`
- `EncryptStream(SessionHandle As Long, StreamName As String, UnencryptedStream As IUnknown, EncryptedStream As IUnknown)`
- `DecryptStream(SessionHandle As Long, StreamName As String, EncryptedStream As IUnknown, UnencryptedStream As IUnknown)`
- `ShowSettings(SessionHandle As Long, ParentWindow As IUnknown, ReadOnly As Boolean, Remove As Boolean)`
