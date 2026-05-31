# AddIns2

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244B5-0000-0000-C000-000000000046}  

A collection of AddIn objects that represent all the add-ins that are currently available or open in Microsoft Excel, regardless of whether they are installed.

**Remarks:** The contents of the AddIns2 collection correspond to the list of add-ins displayed in the Add-Ins dialog box (Add-Ins command on the Developer tab) and any add-ins that are currently open.

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
  Adds a new add-in to the list of add-ins.
    - `Filename As String` (required): The name of the file that contains the add-in to add to the list in the Add-Ins dialog box.
    - `CopyFile As Variant` (optional): If the add-in file is on a removable medium, specifies whether to copy the add-in to the local hard disk. Specify True to copy the add-in to your hard disk. Specify False to keep the add-in on the removable medium. If this argument is omitted, Microsoft Excel displays a dialog box and asks the user to choose whether to copy the add-in file. This parameter is ignored if the add-in file is already on the hard disk.
