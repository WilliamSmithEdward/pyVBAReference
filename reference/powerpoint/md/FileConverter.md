# FileConverter

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {92D41A51-F07E-4CA4-AF6F-BEF486AA4E6F}  

Represents a file converter that is used to open or save files. The FileConverter object is a member of the FileConverters collection. The FileConverters collection contains all the installed file converters for opening and saving files.

**Remarks:** Use FileConverters (Index), where Index is a class name or index number, to return a single FileConverter object. The following example displays the extensions associated with the Microsoft Excel worksheet converter. The index number represents the position of the file converter in the FileConverters collection. The following example displays the format name of the first file converter. You cannot create a new file converter or add one to the FileConverters collection. FileConverter objects are added during installation of Microsoft Office or by installing supplemental file converters. Use either the CanSave or CanOpen property to determine whether a FileConverter object can be used to open or save document. File converters for saving documents are listed in the Save As dialog box. File converters for opening documents appear in a dialog box if the Confirm conversion at Open check box is selected on the General tab in the Options dialog box.

## Properties (12)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft PowerPoint application.
- `CanOpen As Boolean  (read-only)`  
  True if the specified file converter is designed to open files. Read-only Boolean.
- `CanSave As Boolean  (read-only)`  
  True if the specified file converter is designed to save files. Read-only Boolean.
- `ClassName As String  (read-only)`  
  Returns a unique name that identifies the file converter. Read-only String.
- `Creator As FileConverters  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Extensions As String  (read-only)`  
  Returns the file name extensions associated with the specified FileConverter object. Read-only String.
- `FormatName As String  (read-only)`  
  Returns the name of the specified file converter. Read-only String.
- `Name As String  (read-only)`  
  Returns the name of the specified object. Read-only String.
- `Parent As FileConverters  (read-only)`  
  Returns an Object that represents the parent object of the specified FileConverter object.
- `Path As String  (read-only)`  
  Returns the disk or Web path to the specified object. Read-only String.
- `OpenFormat As Long  (read-only)`  
  Returns the file format of the specified file converter. Read-only Long.
- `SaveFormat As Long  (read-only)`  
  Returns the file format of the specified document or file converter. Read-only Long.
