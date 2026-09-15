# AddIn

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {91493461-5A91-11CF-8700-00AA0060263B}  

Represents a single add-in, either loaded or not loaded.

**Remarks:** The AddIn object is a member of the AddIns collection. The AddIns collection contains all of the Microsoft PowerPoint-specific add-ins available, regardless of whether or not they are loaded. The collection does not include Component Object Model (COM) add-ins.

**Example:**

```vba
AddIns("my ppt tools").Loaded = True
```

## Properties (8)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `FullName As String  (read-only)`  
  Returns the name of the specified add-in or saved presentation, including the path, the current file system separator, and the file name extension. Read-only.
- `Name As String  (read-only)`  
  The name (title) of the add-in for file types that are registered. Read-only.
- `Path As String  (read-only)`  
  Returns a String that represents the path to the specified AddIn object. Read-only.
- `Registered As MsoTriState  (read/write)`  
  Determines whether the specified add-in is registered in the Windows registry. Read/write.
- `AutoLoad As MsoTriState  (read/write)`  
  Determines whether the specified add-in is automatically loaded each time PowerPoint is started. Read/write.
- `Loaded As MsoTriState  (read/write)`  
  Determines whether the specified add-in is loaded. Read/write.
