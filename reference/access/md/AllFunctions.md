# AllFunctions

**Type:** Class  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {08F6C822-3CFD-11D1-98BC-006008197D41}  

The AllFunctions collection contains an AccessObject object for each function in the CurrentData or CodeData object.

**Remarks:** The CurrentData or CodeData object has an AllFunctions collection containing AccessObject objects that describe instances of all functions specified by the CurrentData or CodeData objects. For example, you can enumerate the AllFunctions collection in Visual Basic to set or return the values of properties of individual AccessObject objects in the collection. Refer to an individual AccessObject object in the AllFunctions collection either by referring to the object by name, or by referring to its index within the collection. If you want to refer to a specific object in the AllFunctions collection, it's better to refer to the function by name because a function's collection index may change. The AllFunctions collection is indexed beginning with zero. If you refer to a function by its index, the first function is AllFunctions(0), the second function is AllFunctions(1), and so on. To list all open functions in the database, use the IsLoaded property of each AccessObject object in the AllFunctions collection. You can then use the Name property of each individual AccessObject object to return the name of a function.

## Properties (4)

- `Application As Application  (read-only)`  
  Use the Application property to access the active Microsoft Access Application object and its related properties. Read-only Application object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Item As AccessObject  (read-only)`  
  The Item property returns a specific member of a collection either by position or by index. Read-only AccessObject.
- `Count As Long  (read-only)`  
  Use the Count property to determine the number of items in a specified collection. Read-only Long.
