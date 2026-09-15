# AllMacros

**Type:** Class  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {08F6C815-3CFD-11D1-98BC-006008197D41}  

The AllMacros collection contains an AccessObject for each macro in the CurrentProject or CodeProject object.

**Remarks:** The CurrentProject or CodeProject object has an AllMacros collection containing AccessObject objects that describe instances of all the macros specified by CurrentProject or CodeProject. For example, you can enumerate the AllMacros collection in Visual Basic to set or return the values of properties of individual AccessObject objects in the collection. Refer to an individual AccessObject object in the AllMacros collection either by referring to the object by name, or by referring to its index within the collection. If you want to refer to a specific object in the AllMacros collection, it's better to refer to the macro by name because a macro's collection index may change. The AllMacros collection is indexed beginning with zero. If you refer to a macro by its index, the first macro is AllMacros(0), the second macro is AllMacros(1), and so on. You can't add or delete an AccessObject object from the AllMacros collection.

**Example:**

```vba
Sub AllMacros()
 Dim obj As AccessObject, dbs As Object
 Set dbs = Application.CurrentProject
 ' Search for open AccessObject objects in AllMacros collection.
 For Each obj In dbs.AllMacros
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
