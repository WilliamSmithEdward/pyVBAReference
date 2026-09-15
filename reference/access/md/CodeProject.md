# CodeProject

**Type:** Class  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {9212BA72-3E79-11D1-98BD-006008197D41}  

The CodeProject object refers to the project for the code database of a Microsoft Access project (.adp) or Access database.

**Remarks:** The CodeProject object has several collections that contain specific AccessObject objects within the code database. The following table lists the name of each collection defined by Access project and the types of objects it contains. For example, an AccessObject object representing a form is a member of the AllForms collection, which is a collection of AccessObject objects within the current database. Within the AllForms collection, individual members of the collection are indexed beginning with zero. Refer to an individual AccessObject object in the AllForms collection either by referring to the form by name, or by referring to its index within the collection. If you want to refer to a specific object in the AllForms collection, it's better to refer to it by name because an item's collection index may change. If the object name includes a space, the name must be surrounded by brackets ([ ]).

## Properties (23)

- `AllForms As AllForms  (read-only)`  
  Use the AllForms property to reference the AllForms collection and its related properties. Read-only AllForms object.
- `AllReports As AllReports  (read-only)`  
  Use the AllReports property to reference the AllReports collection and its related properties. Read-only AllReports object.
- `AllMacros As AllMacros  (read-only)`  
  Use the AllMacros property to reference the AllMacros collection and its related properties. Read-only AllMacros object.
- `AllModules As AllModules  (read-only)`  
  Use the AllModules property to reference the AllModules collection and its related properties. Read-only AllModules object.
- `ProjectType As AcProjectType  (read-only)`  
  Use the ProjectType property to determine the type of project that is currently open. Read-only AcProjectType.
- `BaseConnectionString As String  (read-only)`  
  Use the BaseConnectionString property to return the base connection string for the specified object. Read-only String.
- `IsConnected As Boolean  (read-only)`  
  Use the IsConnected property to determine if the CodeProject object is currently connected. Read-only Boolean.
- `Name As String  (read-only)`  
  Use the Name property to determine the string expression that identifies the name of an object. Read-only String.
- `Path As String  (read-only)`  
  Use the Path property to determine the location where data is stored for a Microsoft Access project (.adp) or Access database. Read-only String.
- `FullName As String  (read-only)`  
  Sets or returns the full path (including the file name) of a specific object. Read-only String.
- `Connection As Connection  (read-only)`  
  Use the Connection property to return a reference to the current ActiveX Data Objects (ADO) Connection object and its related properties. Read-only Connection.
- `Properties As AccessObjectProperties  (read-only)`  
  Returns a reference to a CodeProject object's AccessObjectProperties collection. Read-only.
- `Application As Application  (read-only)`  
  Use the Application property to access the active Microsoft Access Application object and its related properties. Read-only Application object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `RemovePersonalInformation As Boolean  (read/write)`  
  Returns or sets a Boolean indicating whether personal information about the user is stored in the specified project. True if personal information is removed. Read/write.
- `FileFormat As AcFileFormat  (read-only)`  
  Returns an AcFileFormat constant indicating the Microsoft Access version format of the specified project. Read-only.
- `AccessConnection As Connection  (read-only)`  
  Use the AccessConnection property to return a reference to the current Microsoft ActiveX Data Objects (ADO) Connection object and its related properties. Read-only Connection.
- `ImportExportSpecifications As ImportExportSpecifications  (read-only)`  
  Returns an ImportExportSpecifications collection that represents the collection of saved import or export operations for the specified object. Read-only.
- `IsTrusted As Boolean  (read-only)`  
  Gets whether macros and Visual Basic for Applications (VBA) code have been enabled in the current project. Read-only Boolean.
- `WebSite As String  (read-only)`  
  Gets the Uniform Resource Locator (URL) of the website to which the database has been published. Read-only String.
- `IsWeb As Boolean  (read-only)`  
  Gets whether the database is a web database. Read-only Boolean.
- `Resources As SharedResources  (read-only)`  
  Gets the SharedResources collection for the specified object. Read-only SharedResources.
- `IsSQLBackend As Boolean  (read-only)`  
  Returns True if the code project was created in Access 2013 and onwards and False if the code project was created prior to Access 2013 . Read-only Boolean Introduced in Office 2016.

## Methods (4)

- `OpenConnection([BaseConnectionString As Variant], [UserID As Variant], [Password As Variant])`  
  Use the OpenConnection method to open an ADO connection to an existing Microsoft Access project (.adp) or Access database as the current Access project or database in the Microsoft Access window.
    - `BaseConnectionString As Variant` (optional): A string expression that is the base connection string of the database.
    - `UserID As Variant` (optional): A string expression that is the name of the existing Access project, including the path name and the file name extension. If your network supports it, you can also specify a network path in the following form: \\Server\Share\Folder\Filename.adp
    - `Password As Variant` (optional): If you don't supply the file name extension, .adp is appended to the filename. Use this method or the OpenCurrentDatabase method to open .adp files.
- `CloseConnection()`  
  Use the CloseConnection method to close the current connection between the CodeProject object in a Microsoft Access project (.adp) or Access database and the database specified in the project's base connection string.
- `UpdateDependencyInfo()`  
  Updates the dependency information for the database.
- `AddSharedImage(SharedImageName As String, FileName As String)`  
  Imports the specified image into the database and adds it to the SharedResources collection.
    - `SharedImageName As String` (required): Specifies the string used to identify the image in the collection.
    - `FileName As String` (required): Specifies the full name and path to the image file.
