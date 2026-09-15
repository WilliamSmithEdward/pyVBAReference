# Modules

**Type:** Dispatch Interface  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {9DD0AF42-6E28-11CF-9008-00AA0042B7CE}  

The Modules collection contains all open standard modules and class modules in a Microsoft Access database.

**Remarks:** All open modules are included in the Modules collection, whether they are uncompiled, compiled, or in break mode, or contain the code that's running. To determine whether an individual Module object represents a standard module or a class module, check the Module object's Type property. The Modules collection belongs to the Microsoft Access Application object. Individual Module objects in the Modules collection are indexed beginning with zero.

**Example:**

```vba
Sub PrintOpenModuleNames()
 Dim i As Integer
 Dim modOpenModules As Modules

 Set modOpenModules = Application.Modules

 For i = 0 To modOpenModules.Count - 1

 Debug.Print modOpenModules(i).Name

 Next
End Sub
```

## Properties (4)

- `Application As Application  (read-only)`  
  Use the Application property to access the active Microsoft Access Application object and its related properties. Read-only Application object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Item As Module  (read-only)`  
  The Item property returns a specific member of a collection either by position or by index. Read-only Module.
- `Count As Long  (read-only)`  
  Use the Count property to determine the number of items in a specified collection. Read-only Long.
