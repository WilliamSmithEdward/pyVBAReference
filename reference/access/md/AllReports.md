# AllReports

**Type:** Class  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {08F6C814-3CFD-11D1-98BC-006008197D41}  

The AllReports collection contains an AccessObject for each report in the CurrentProject or CodeProject object.

**Remarks:** The CurrentProject or CodeProject object has an AllReports collection containing AccessObject objects that describe instances of all the reports in the database. For example, you can enumerate the AllReports collection in Visual Basic to set or return the values of properties of individual AccessObject objects in the collection. Refer to an individual AccessObject object in the AllReports collection either by referring to the item by name, or by referring to its index within the collection. If you want to refer to a specific report in the AllReports collection, it's better to refer to the report by name because the index may change. The AllReports collection is indexed beginning with zero. If you refer to a report by its index, the first report is AllReports(0), the second report is AllReports(1), and so on. You can't add or delete an AccessObject object from the AllReports collection.

**Example:**

```vba
Sub AllReports()
 Dim obj As AccessObject, dbs As Object
 Set dbs = Application.CurrentProject
 ' Search for open AccessObject objects in AllReports collection.
 For Each obj In dbs.AllReports
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
