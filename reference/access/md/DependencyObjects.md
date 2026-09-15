# DependencyObjects

**Type:** Class  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {1C4367A8-EAEE-4C23-9582-4A229DF2403E}  

Represents the collection of AccessObject objects that a DependencyInfo object contains.

**Remarks:** To return a DependencyObjects collection, use the Dependants or Dependencies property of the DependencyInfo object. To return a single AccessObject object, use the Item property of the DependencyObjects collection.

## Properties (4)

- `Application As Application  (read-only)`  
  Use the Application property to access the active Microsoft Access Application object and its related properties. Read-only Application object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Item As AccessObject  (read-only)`  
  The Item property returns a specific member of a collection either by position or by index. Read-only AccessObject.
- `Count As Long  (read-only)`  
  Use the Count property to determine the number of items in a specified collection. Read-only Long.
