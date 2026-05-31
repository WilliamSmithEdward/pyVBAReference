# Dialogs

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00020879-0000-0000-C000-000000000046}  

A collection of all the Dialog objects in Microsoft Excel.

**Remarks:** Each Dialog object represents a built-in dialog box. You cannot create a new built-in dialog box or add one to the collection. The only useful thing that you can do with a Dialog object is to use it with the Show method to display the corresponding dialog box. The Microsoft Excel Visual Basic object library includes built-in constants for many of the built-in dialog boxes. Each constant is formed from the prefix "xlDialog" followed by the name of the dialog box. For example, the Apply Names dialog box constant is xlDialogApplyNames, and the Find File dialog box constant is xlDialogFindFile. These constants are members of the XlBuiltinDialog enumerated type.

**Example:**

```vba
MsgBox Application.Dialogs.Count
```

## Properties (7)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Count As Long  (read-only)`  
  Returns a Long value that represents the number of objects in the collection.
- `Item As Dialog  (read-only)`  
  Returns a single object from a collection.
- `_Default As Dialog  (read-only)`
- `_NewEnum As IUnknown  (read-only)`
