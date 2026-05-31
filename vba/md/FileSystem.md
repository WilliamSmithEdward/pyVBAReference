# FileSystem

**Type:** Module  
**Library:** Visual Basic For Applications  

## Functions (20)

- `ChDir(Path As String)`  
  Changes the current directory or folder.
- `ChDrive(Drive As String)`  
  Changes the current drive.
- `EOF(FileNumber As Integer) As Boolean`  
  Returns an Integer containing the Boolean value True when the end of a file opened for Random or sequential Input has been reached.
- `FileAttr(FileNumber As Integer, [ReturnType As Integer]) As Long`  
  Returns a Long representing the file mode for files opened by using the Open statement.
    - `FileNumber As Integer` (required): Required; Integer. Any valid file number.
    - `ReturnType As Integer` (optional): Required; Integer. Number indicating the type of information to return. Specify 1 to return a value indicating the file mode. On 16-bit systems only, specify 2 to retrieve an operating system file handle. _Returntype_ 2 is not supported in 32-bit systems and causes an error.
- `FileCopy(Source As String, Destination As String)`  
  Copies a file.
    - `Source As String` (required): Required. String expression that specifies the name of the file to be copied. The _source_ may include directory or folder, and drive.
    - `Destination As String` (required): Required. String expression that specifies the target file name. The _destination_ may include directory or folder, and drive.
- `FileDateTime(PathName As String) As Variant`  
  Returns a Variant (Date) that indicates the date and time when a file was created or last modified.
- `FileLen(PathName As String) As Long`  
  Returns a Long specifying the length of a file in bytes.
- `GetAttr(PathName As String) As VbFileAttribute`  
  Returns an Integer representing the attributes of a file, directory, or folder.
- `Kill(PathName As Variant)`  
  Deletes files from a disk.
- `Loc(FileNumber As Integer) As Long`  
  Returns a Long specifying the current read/write position within an open file.
- `LOF(FileNumber As Integer) As Long`  
  Returns a Long representing the size, in bytes, of a file opened by using the Open statement.
- `MkDir(Path As String)`  
  Creates a new directory or folder.
- `Reset()`  
  Closes all disk files opened by using the Open statement.
- `RmDir(Path As String)`  
  Removes an existing directory or folder.
- `Seek(FileNumber As Integer) As Long`  
  Returns a Long specifying the current read/write position within a file opened by using the Open statement.
- `SetAttr(PathName As String, Attributes As VbFileAttribute)`  
  Sets attribute information for a file.
    - `PathName As String` (required): Required. String expression that specifies a file name; may include directory or folder, and drive.
    - `Attributes As VbFileAttribute` (required): Required. Constant or numeric expression whose sum specifies file attributes.
- `_B_str_CurDir([Drive As Variant]) As String`
- `_B_var_CurDir([Drive As Variant]) As Variant`
- `FreeFile([RangeNumber As Variant]) As Integer`  
  Returns an Integer representing the next file number available for use by the Open statement.
- `Dir([PathName As Variant], [Attributes As VbFileAttribute]) As String`  
  Returns a String representing the name of a file, directory, or folder that matches a specified pattern or file attribute, or the volume label of a drive.
    - `PathName As Variant` (optional): Optional. String expression that specifies a file name; may include directory or folder, and drive. A zero-length string ("") is returned if _pathname_ is not found.
    - `Attributes As VbFileAttribute` (optional): Optional. Constant or numeric expression, whose sum specifies file attributes. If omitted, returns files that match _pathname_ but have no attributes.
