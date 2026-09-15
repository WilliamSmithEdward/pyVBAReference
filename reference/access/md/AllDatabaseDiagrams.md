# AllDatabaseDiagrams

**Type:** Class  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {08F6C821-3CFD-11D1-98BC-006008197D41}  

The AllDatabaseDiagrams collection contains an AccessObject for each database diagram in the CurrentData or CodeData object.

**Remarks:** The CurrentData or CodeData object has an AllDatabaseDiagrams collection containing AccessObject objects that describe instances of all database diagrams specified by CurrentData or CodeData. For example, you can enumerate the AllDatabaseDiagrams collection in Visual Basic to set or return the values of properties of individual AccessObject objects in the collection. Refer to an individual AccessObject object in the AllDatabaseDiagrams collection either by referring to the object by name, or by referring to its index within the collection. If you want to refer to a specific object in the AllDatabaseDiagrams collection, it's better to refer to the database diagram by name because a database diagram's collection index may change. The AllDatabaseDiagrams collection is indexed beginning with zero. If you refer to a database diagram by its index, the first database diagram is AllDatabaseDiagrams(0), the second database diagram is AllDatabaseDiagrams(1), and so on.

**Example:**

```vba
Sub AllDatabaseDiagrams()
    Dim obj As AccessObject, dbs As Object
    Set dbs = Application.CurrentData
    ' Search for open AccessObject objects in
    ' AllDatabaseDiagrams collection.
    For Each obj In dbs.AllDatabaseDiagrams
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
