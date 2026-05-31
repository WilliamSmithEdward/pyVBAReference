# EncryptionProvider

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000CD809-0000-0000-C000-000000000046}  

Provides the methods for setting up permissions, applying the cryptography of the underlying encryption and decryption, and user authentication.

**Remarks:** Encryption providers are implemented through custom COM add-ins. You are provided with storage within Office documents for add-in specific information to store whatever information you need to encrypt, decrypt, apply rights, and display permission setup or authentication user interfaces.

## Methods (9)

- `GetProviderDetail(encprovdet As EncryptionProviderDetail) As Variant`  
  Displays information about the encryption of the current document.
    - `encprovdet As EncryptionProviderDetail` (required): Specifies the encryption information that you want.
- `NewSession(ParentWindow As IUnknown) As Long`  
  Used by the EncryptionProvider object to create a new encryption session. This session is used by the provider to cache document-specific information about the encryption, users, and rights while the document is in memory.
    - `ParentWindow As IUnknown` (required): Specifies the window that is called to display the encryption settings.
- `Authenticate(ParentWindow As IUnknown, EncryptionData As IUnknown, PermissionsMask As Long) As Long`  
  Used to determine whether the user has the proper permissions to open the encrypted document.
    - `ParentWindow As IUnknown` (required): Specifies the window that is called to display the encryption settings.
    - `EncryptionData As IUnknown` (required): Contains the encrypted data for the current document.
    - `PermissionsMask As Long` (required): The user interface displayed by the encryption provider add-in.
- `CloneSession(SessionHandle As Long) As Long`  
  Creates a second, working copy of the EncryptionProvider object's encryption session for a file that is about to be saved.
    - `SessionHandle As Long` (required): The ID of the cloned session.
- `EndSession(SessionHandle As Long)`  
  Ends the current encryption session.
    - `SessionHandle As Long` (required): The ID of the current session.
- `Save(SessionHandle As Long, EncryptionData As IUnknown) As Long`  
  Saves an encrypted document.
    - `SessionHandle As Long` (required): The ID of the current session.
    - `EncryptionData As IUnknown` (required): Contains the encryption information.
- `EncryptStream(SessionHandle As Long, StreamName As String, UnencryptedStream As IUnknown, EncryptedStream As IUnknown)`  
  Encrypts and returns a stream of data for a document.
    - `SessionHandle As Long` (required): The ID of the current session.
    - `StreamName As String` (required): The name of the encrypted stream of document data.
    - `UnencryptedStream As IUnknown` (required): The data stream before encryption.
    - `EncryptedStream As IUnknown` (required): The data stream information after it has been encrypted.
- `DecryptStream(SessionHandle As Long, StreamName As String, EncryptedStream As IUnknown, UnencryptedStream As IUnknown)`  
  Decrypts and returns a stream of encrypted data for a document.
    - `SessionHandle As Long` (required): The ID of the current session.
    - `StreamName As String` (required): The ID of the stream of data.
    - `EncryptedStream As IUnknown` (required): The encrypted data stream.
    - `UnencryptedStream As IUnknown` (required): The data stream before dencryption.
- `ShowSettings(SessionHandle As Long, ParentWindow As IUnknown, ReadOnly As Boolean, Remove As Boolean)`  
  Used to display a dialog of the encryption settings for the current document.
    - `SessionHandle As Long` (required): The ID of the current session.
    - `ParentWindow As IUnknown` (required): Specifies the window that is called to display the encryption settings.
    - `ReadOnly As Boolean` (required): Specifies whether you want the user to be able to change the encryption settings.
    - `Remove As Boolean` (required): If True, the encryption for a document will be removed during the next save operation.
