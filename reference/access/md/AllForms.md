# AllForms

**Type:** Class  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {08F6C813-3CFD-11D1-98BC-006008197D41}  

The AllForms collection contains an AccessObject object for each form in the CurrentProject or CodeProject object.

**Remarks:** The CurrentProject and CodeProject object has an AllForms collection containing AccessObject objects that describe instances of all the forms in the database. For example, you can enumerate the AllForms collection in Visual Basic to set or return the values of properties of individual AccessObject objects in the collection. Refer to an individual AccessObject object in the AllForms collection either by referring to the object by name, or by referring to its index within the collection. If you want to refer to a specific object in the AllForms collection, it's better to refer to the form by name because a form's collection index may change. The AllForms collection is indexed beginning with zero. If you refer to a form by its index, the first form is AllForms(0), the second form is AllForms(1), and so on. You can't add or delete an AccessObject object from the AllForms collection.

**Example:**

```vba
Sub AllForms()
    Dim obj As AccessObject, dbs As Object
    Set dbs = Application.CurrentProject
    ' Search for open AccessObject objects in AllForms collection.
    For Each obj In dbs.AllForms
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
