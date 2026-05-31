# IFolder

**Type:** Dispatch Interface  
**Library:** Microsoft Scripting Runtime  
**GUID:** {C7C3F5A2-88A3-11D0-ABCB-00A0C90FFFC0}  

Folder Interface

## Properties (15)

- `Path As String  (read-only)`  
  Path to folder
- `Name As String  (read/write)`  
  Get name of folder
- `ShortPath As String  (read-only)`  
  Short path
- `ShortName As String  (read-only)`  
  Short name
- `Drive As IDrive  (read-only)`  
  Get drive that contains folder
- `ParentFolder As IFolder  (read-only)`  
  Get parent folder
- `Attributes As FileAttribute  (read/write)`  
  Folder attributes
- `DateCreated As Date  (read-only)`  
  Date folder was created
- `DateLastModified As Date  (read-only)`  
  Date folder was last modified
- `DateLastAccessed As Date  (read-only)`  
  Date folder was last accessed
- `Type As String  (read-only)`  
  Type description
- `IsRootFolder As Boolean  (read-only)`  
  True if folder is root
- `Size As Variant  (read-only)`  
  Sum of files and subfolders
- `SubFolders As IFolderCollection  (read-only)`  
  Get folders collection
- `Files As IFileCollection  (read-only)`  
  Get files collection

## Methods (4)

- `Delete([Force As Boolean])`  
  Delete this folder
- `Copy(Destination As String, [OverWriteFiles As Boolean])`  
  Copy this folder
- `Move(Destination As String)`  
  Move this folder
- `CreateTextFile(FileName As String, [Overwrite As Boolean], [Unicode As Boolean]) As ITextStream`  
  Create a file as a TextStream
