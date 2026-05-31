# AddIns

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00020858-0000-0000-C000-000000000046}  

A collection of AddIn objects that represents all the add-ins available to Microsoft Excel, regardless of whether they're installed.

**Remarks:** This list corresponds to the list of add-ins displayed in the Add-Ins dialog box.

**Example:**

```vba
Sub DisplayAddIns()
 Worksheets("Sheet1").Activate
 rw = 1
 For Each ad In Application.AddIns
 Worksheets("Sheet1").Cells(rw, 1) = ad.Name
 Worksheets("Sheet1").Cells(rw, 2) = ad.Installed
 rw = rw + 1
 Next
End Sub
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
- `Item As AddIn  (read-only)`  
  Returns a single object from a collection.
- `_NewEnum As IUnknown  (read-only)`
- `_Default As AddIn  (read-only)`

## Methods (1)

- `Add(Filename As String, [CopyFile As Variant]) As AddIn`  
  Adds a new add-in file to the list of add-ins. Returns an AddIn object.
    - `Filename As String` (required): The name of the file that contains the add-in or the ProgID of the automation add-in that you want to add to the list in the add-in manager.
    - `CopyFile As Variant` (optional): Ignored if the add-in file is on a hard disk. True to copy the add-in to your hard disk, if the add-in is on a removable medium (such as a compact disc). False to have the add-in remain on the removable medium. If this argument is omitted, Microsoft Excel displays a dialog box and asks you to choose.
