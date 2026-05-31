# AddIn

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00020857-0000-0000-C000-000000000046}  

Represents a single add-in, either installed or not installed.

**Remarks:** The AddIn object is a member of the AddIns collection. The AddIns collection contains a list of all the add-ins available to Microsoft Excel, regardless of whether they're installed. This list corresponds to the list of add-ins displayed in the Add-Ins dialog box.

**Example:**

```vba
AddIns("analysis toolpak").Installed = True
```

## Properties (10)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `FullName As String  (read-only)`  
  Returns the name of the object, including its path on disk, as a string. Read-only String.
- `Installed As Boolean  (read/write)`  
  True if the add-in is installed or to install the add-in; False if the add-in is uninstalled or to uninstall the add-in. Read/write Boolean.
- `Name As String  (read-only)`  
  Returns a String value that represents the name of the object.
- `Path As String  (read-only)`  
  Returns a String value that represents the complete path to the Add-in, excluding the final separator and name of the Add-in.
- `progID As String  (read-only)`  
  Returns the programmatic identifiers for the object. Read-only String.
- `CLSID As String  (read-only)`  
  Returns a read-only unique identifier, or CLSID, identifying an object as a String.
- `IsOpen As Boolean  (read-only)`  
  Returns True if the add-in is currently open. Read-only Boolean.
