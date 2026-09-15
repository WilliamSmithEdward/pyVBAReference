# AccessObject

**Type:** Dispatch Interface  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {ABE316B1-3FF6-11D1-98BD-006008197D41}  

An AccessObject object refers to a particular Access object.

**Remarks:** An AccessObject object includes information about one instance of an object. The following table list the types of objects each AccessObject describes, the name of its collection, and what type of information AccessObject contains. Because an AccessObject object corresponds to an existing object, you can't create new AccessObject objects or delete existing ones. To refer to an AccessObject object in a collection by its ordinal number or by its Name property setting, use any of the following syntax forms: - AllForms (0) - AllForms ("name") - AllForms ![ name ]

## Properties (10)

- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Name As String  (read-only)`  
  Use the Name property to determine the string expression that identifies the name of an object. Read-only String.
- `Type As AcObjectType  (read-only)`  
  Returns the value of an AccessObject object type. Read-only AcObjectType.
- `Properties As AccessObjectProperties  (read-only)`  
  Returns a reference to an AccessObject object's AccessObjectProperties collection. Read-only.
- `IsLoaded As Boolean  (read-only)`  
  Use the IsLoaded property to determine if an AccessObject object is currently loaded. Read-only Boolean.
- `FullName As String  (read/write)`  
  Sets or returns the full path (including the file name) of a specific object. Read/write String.
- `DateCreated As Date  (read-only)`  
  Returns a Date indicating the date and time when the design of the specified object was last modified. Read-only.
- `DateModified As Date  (read-only)`  
  Returns a Date indicating the date and time when the design of the specified object was last modified. Read-only.
- `CurrentView As AcCurrentView  (read-only)`  
  Returns the current view for the specified Access object. Read-only AcCurrentView.
- `IsWeb As Boolean  (read-only)`  
  Gets whether the specified object is a web object. Read-only Boolean.

## Methods (2)

- `IsDependentUpon(ObjectType As AcObjectType, ObjectName As String) As Boolean`  
  Returns a Boolean value that indicates whether the specified object is dependent upon the database object specified in the ObjectName argument.
    - `ObjectType As AcObjectType` (required): An AcObjectType constant that represents the type of database object to check for dependency.
    - `ObjectName As String` (required): The name of the database object to check for dependency.
- `GetDependencyInfo() As _DependencyInfo`  
  Returns a DependencyInfo object that represents the database objects that are dependent upon the specified object.
