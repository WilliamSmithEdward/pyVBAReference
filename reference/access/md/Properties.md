# Properties

**Type:** Dispatch Interface  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {331FDD02-CF31-11CD-8701-00AA003F0F07}  

The Properties collection contains all the built-in properties in an instance of an open form, report, or control. These properties uniquely characterize that instance of the object.

**Remarks:** Use the Properties collection in Visual Basic or in an expression to refer to form, report, or control properties on forms or reports that are currently open. Use the Properties collection of an object to enumerate the object's built-in properties. You don't need to know beforehand exactly which properties exist or what their characteristics (Name and Value properties) are to manipulate them.

**Example:**

```vba
Sub AllOpenForms()
 Dim frm As Form, prp As Property

 ' Enumerate Forms collection.
 For Each frm In Forms
 ' Print name of form.
 Debug.Print frm.Name
 ' Enumerate Properties collection of each form.
 For Each prp In frm.Properties
 ' Print name of each property.
 Debug.Print prp.Name; " = "; prp.Value
 Next prp
 Next frm
End Sub
```

## Properties (4)

- `Application As Application  (read-only)`  
  Use the Application property to access the active Microsoft Access Application object and its related properties. Read-only Application object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Item As Object  (read-only)`  
  The Item property returns a specific member of a collection either by position or by index. Read-only Object.
- `Count As Long  (read-only)`  
  Returns the number of objects in the specified collection. Read-only.
