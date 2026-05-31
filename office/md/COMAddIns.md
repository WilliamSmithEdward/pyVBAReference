# COMAddIns

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C0339-0000-0000-C000-000000000046}  

A collection of COMAddIn objects that provide information about a COM add-in registered in the Windows registry.

**Example:**

```vba
MsgBox Application.COMAddIns.Count
```

## Properties (4)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the COMAddIns object (you can use this property with an Automation object to return that object's container application). Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the COMAddIns object was created. Read-only.
- `Count As Long  (read-only)`  
  Gets a count of the number of COM add-ins in the host application. Read-only.
- `Parent As Object  (read-only)`  
  Gets the Parent object for the COMAddIns object. Read-only.

## Methods (2)

- `Item(Index As Variant) As COMAddIn`  
  Gets a member of the specified COMAddIns collection.
    - `Index As Variant` (required): Represents the location of the member within the collection.
- `Update()`  
  Updates the contents of the COMAddIns collection from the list of add-ins stored in the Windows registry.
