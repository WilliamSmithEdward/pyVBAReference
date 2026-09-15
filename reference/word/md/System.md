# System

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020935-0000-0000-C000-000000000046}  

Contains information about the computer system.

**Remarks:** Use the System property to return the System object. If the operating system is Windows, the following example makes a network connection to \\Project\Info. The following example displays the current screen resolution (for example, "1024 x 768").

## Properties (17)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified System object.
- `OperatingSystem As String  (read-only)`  
  Returns the name of the current operating system (for example, "Windows" or "Windows Server"). Read-only String.
- `Version As String  (read-only)`  
  Returns the version number of the operating system. Read-only String.
- `FreeDiskSpace As Long  (read-only)`  
  Returns the available disk space for the current drive, in bytes. Use the ChDrive statement to change the current drive. Read-only Long.
- `LanguageDesignation As String  (read-only)`  
  Returns the designated language of the system software. Read-only String.
- `HorizontalResolution As Long  (read-only)`  
  Returns the horizontal display resolution, in pixels. Read-only Long.
- `VerticalResolution As Long  (read-only)`  
  Returns the vertical screen resolution in pixels. Read-only Long.
- `ProfileString As String  (read/write)`  
  Returns or sets a value for an entry in the Windows registry under the following subkey: HKEY_CURRENT_USER\Software\Microsoft\Office\version\Word. Read/write String.
- `PrivateProfileString As String  (read/write)`  
  Returns or sets a string in a settings file or the Microsoft Windows registry. Read/write String.
- `MathCoprocessorInstalled As Boolean  (read-only)`  
  True if a math coprocessor is installed on the system. Read-only Boolean.
- `ComputerType As String  (read-only)`
- `MacintoshName As String  (read-only)`
- `QuickDrawInstalled As Boolean  (read-only)`
- `Cursor As WdCursorType  (read/write)`  
  Returns or sets the state (shape) of the pointer. Can be one of the following WdCursorType constants: wdCursorIBeam, wdCursorNormal, wdCursorNorthwestArrow, or wdCursorWait. Read/write Long.
- `CountryRegion As WdCountry  (read-only)`  
  Returns the country/region designation of the system. Read-only WdCountry.

## Methods (2)

- `MSInfo()`  
  Starts the Microsoft System Information application if it is not running, or switches to it if it is already running.
- `Connect(Path As String, [Drive As Variant], [Password As Variant])`  
  Establishes a connection to a network drive.
    - `Path As String` (required): The path for the network drive (for example, "\\Project\Info").
    - `Drive As Variant` (optional): A number corresponding to the letter you want to assign to the network drive, where 0 (zero) corresponds to the first available drive letter, 1 corresponds to the second available drive letter, and so on. If this argument is omitted, the next available letter is used.
    - `Password As Variant` (optional): The password, if the network drive is protected with a password.
