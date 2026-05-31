# IDrive

**Type:** Dispatch Interface  
**Library:** Microsoft Scripting Runtime  
**GUID:** {C7C3F5A0-88A3-11D0-ABCB-00A0C90FFFC0}  

Drive Interface

## Properties (12)

- `Path As String  (read-only)`  
  Path
- `DriveLetter As String  (read-only)`  
  Drive letter
- `ShareName As String  (read-only)`  
  Share name
- `DriveType As DriveTypeConst  (read-only)`  
  Drive type
- `RootFolder As IFolder  (read-only)`  
  Root folder
- `AvailableSpace As Variant  (read-only)`  
  Get available space
- `FreeSpace As Variant  (read-only)`  
  Get drive free space
- `TotalSize As Variant  (read-only)`  
  Get total drive size
- `VolumeName As String  (read/write)`  
  Name of volume
- `FileSystem As String  (read-only)`  
  Filesystem type
- `SerialNumber As Long  (read-only)`  
  Serial number
- `IsReady As Boolean  (read-only)`  
  Check if disk is available
