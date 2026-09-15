# AllModules

**Type:** Class  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {08F6C816-3CFD-11D1-98BC-006008197D41}  

The AllModules collection contains an AccessObject of each module in the CurrentProject or CodeProject object.

**Remarks:** The CurrentProject or CodeProject object has an AllModules collection containing AccessObject objects that describe instances of all the Module objects specified by CurrentProject or CodeProject. For example, you can enumerate the AllModules collection in Visual Basic to set or return the values of properties of individual AccessObject objects in the collection. Refer to an individual AccessObject object in the AllModules collection either by referring to the object by name, or by referring to its index within the collection. If you want to refer to a specific object in the AllModules collection, it's better to refer to the module by name because a module's collection index may change. The AllModules collection is indexed beginning with zero. If you refer to a module by its index, the first module is AllModules(0), the second module is AllModules(1), and so on. You can't add or delete an AccessObject object from the AllModules collection.

**Example:**

```vba
Sub AllModules()
 Dim obj As AccessObject, dbs As Object
 Set dbs = Application.CurrentProject
 ' Search for open AccessObject objects in AllModules collection.
 For Each obj In dbs.AllModules
 If obj.IsLoaded = True Then
 ' Print name of obj.
 Debug.Print obj.Name
 End If
 Next obj
End Sub
```

## Properties (4)

- `Application As Application  (read-only)`  
  Use the Application property to access the active Microsoft Access Application object and its related properties. Read-only Application object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Item As AccessObject  (read-only)`  
  The Item property returns a specific member of a collection either by position or by index. Read-only AccessObject.
- `Count As Long  (read-only)`  
  Use the Count property to determine the number of items in a specified collection. Read-only Long.
