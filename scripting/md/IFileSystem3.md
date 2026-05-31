# IFileSystem3

**Type:** Dispatch Interface  
**Library:** Microsoft Scripting Runtime  
**GUID:** {2A0B9D10-4B87-11D3-A97A-00104B365C9F}  

FileSystemObject

## Properties (1)

- `Drives As IDriveCollection  (read-only)`  
  Get drives collection

## Methods (26)

- `BuildPath(Path As String, Name As String) As String`  
  Generate a path from an existing path and a name
- `GetDriveName(Path As String) As String`  
  Return drive from a path
- `GetParentFolderName(Path As String) As String`  
  Return path to the parent folder
- `GetFileName(Path As String) As String`  
  Return the file name from a path
- `GetBaseName(Path As String) As String`  
  Return base name from a path
- `GetExtensionName(Path As String) As String`  
  Return extension from path
- `GetAbsolutePathName(Path As String) As String`  
  Return the canonical representation of the path
- `GetTempName() As String`  
  Generate name that can be used to name a temporary file
- `DriveExists(DriveSpec As String) As Boolean`  
  Check if a drive or a share exists
- `FileExists(FileSpec As String) As Boolean`  
  Check if a file exists
- `FolderExists(FolderSpec As String) As Boolean`  
  Check if a path exists
- `GetDrive(DriveSpec As String) As IDrive`  
  Get drive or UNC share
- `GetFile(FilePath As String) As IFile`  
  Get file
- `GetFolder(FolderPath As String) As IFolder`  
  Get folder
- `GetSpecialFolder(SpecialFolder As SpecialFolderConst) As IFolder`  
  Get location of various system folders
- `DeleteFile(FileSpec As String, [Force As Boolean])`  
  Delete a file
- `DeleteFolder(FolderSpec As String, [Force As Boolean])`  
  Delete a folder
- `MoveFile(Source As String, Destination As String)`  
  Move a file
- `MoveFolder(Source As String, Destination As String)`  
  Move a folder
- `CopyFile(Source As String, Destination As String, [OverWriteFiles As Boolean])`  
  Copy a file
- `CopyFolder(Source As String, Destination As String, [OverWriteFiles As Boolean])`  
  Copy a folder
- `CreateFolder(Path As String) As IFolder`  
  Create a folder
- `CreateTextFile(FileName As String, [Overwrite As Boolean], [Unicode As Boolean]) As ITextStream`  
  Create a file as a TextStream
- `OpenTextFile(FileName As String, [IOMode As IOMode], [Create As Boolean], [Format As Tristate]) As ITextStream`  
  Open a file as a TextStream
- `GetStandardStream(StandardStreamType As StandardStreamTypes, [Unicode As Boolean]) As ITextStream`  
  Retrieve the standard input, output or error stream
- `GetFileVersion(FileName As String) As String`  
  Retrieve the file version of the specified file into a string
