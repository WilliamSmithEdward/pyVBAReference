# DependencyInfo

**Type:** Class  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {F8C9DCB3-4063-490E-A73C-3533207CBC26}  

Represents the dependency information for an AccessObject object.

**Remarks:** To return the dependency information for an AccessObject object, use the GetDependencyInfo method of the AccessObject object.

## Properties (6)

- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Dependants As _DependencyObjects  (read-only)`  
  Returns a DependencyObjects collection that represents the objects that depend on the specified AccessObject object. Read-only.
- `Dependencies As _DependencyObjects  (read-only)`  
  Returns a DependencyObjects collection that represents the objects that the specified AccessObject object depends on. Read-only.
- `OutOfDateObjects As _DependencyObjects  (read-only)`  
  Returns a DependencyObjects collection that represents the AccessObject objects for which the dependency information is outdated. Read-only.
- `InsufficientPermissions As _DependencyObjects  (read-only)`  
  Returns a DependencyObjects collection that contains a collection of objects to which you don't have sufficient permissions to check for dependencies. Read-only.
- `UnsupportedObjects As _DependencyObjects  (read-only)`  
  Returns a DependencyObjects collection that contains objects that cannot be searched for dependency. Read-only.
