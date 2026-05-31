# RecentFiles

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024406-0000-0000-C000-000000000046}  

Represents the list of recently used files.

**Remarks:** Each file is represented by a RecentFile object.

**Example:**

```vba
Application.RecentFiles.Maximum = 6
```

## Properties (8)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Maximum As Long  (read/write)`  
  Returns or sets the maximum number of files in the list of recently used files. Can be a value from 0 (zero) through 50. Read/write Long.
- `Count As Long  (read-only)`  
  Returns a Long value that represents the number of objects in the collection.
- `Item As RecentFile  (read-only)`  
  Returns a single object from a collection.
- `_Default As RecentFile  (read-only)`
- `_NewEnum As IUnknown  (read-only)`

## Methods (1)

- `Add(Name As String) As RecentFile`  
  Adds a file to the list of recently used files.
    - `Name As String` (required): The file name.
