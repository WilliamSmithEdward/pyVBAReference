# IFile

**Type:** Dispatch Interface  
**Library:** Microsoft Scripting Runtime  
**GUID:** {C7C3F5A4-88A3-11D0-ABCB-00A0C90FFFC0}  

File Interface

## Properties (12)

- `Path As String  (read-only)`  
  Path to the file
- `Name As String  (read/write)`  
  Get name of file
- `ShortPath As String  (read-only)`  
  Short path
- `ShortName As String  (read-only)`  
  Short name
- `Drive As IDrive  (read-only)`  
  Get drive that contains file
- `ParentFolder As IFolder  (read-only)`  
  Get folder that contains file
- `Attributes As FileAttribute  (read/write)`  
  File attributes
- `DateCreated As Date  (read-only)`  
  Date file was created
- `DateLastModified As Date  (read-only)`  
  Date file was last modified
- `DateLastAccessed As Date  (read-only)`  
  Date file was last accessed
- `Size As Variant  (read-only)`  
  File size
- `Type As String  (read-only)`  
  Type description

## Methods (4)

- `Delete([Force As Boolean])`  
  Delete this file
- `Copy(Destination As String, [OverWriteFiles As Boolean])`  
  Copy this file
- `Move(Destination As String)`  
  Move this file
- `OpenAsTextStream([IOMode As IOMode], [Format As Tristate]) As ITextStream`  
  Open a file as a TextStream
