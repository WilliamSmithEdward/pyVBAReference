# AddIns

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {0002097F-0000-0000-C000-000000000046}  

A collection of AddIn objects that represents all the add-ins available to Word, regardless of whether or not they are currently loaded. The AddIns collection includes global templates or Word add-in libraries (WLLs) displayed in the Templates and Add-ins dialog box.

**Remarks:** Use the AddIns property to return the AddIns collection. The following example displays the name and the installed state of each available add-in. Use the Add method to add an add-in to the list of available add-ins and (optionally) install it using the Install argument. To install an add-in shown in the list of available add-ins, use the Installed property. Use AddIns (index), where index is the add-in name or index number, to return a single AddIn object. You must exactly match the spelling (but not necessarily the capitalization) of the name, as it is shown in the Templates and Add-ins dialog box. To install an add-in shown in the list of available add-ins, use the Installed property. The following example loads the Letter.dot template as a global template. Use the Compiled property to determine whether an AddIn object is a template or a WLL.

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the Addins collection. This is usually an Application object.
- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns the number of AddIn objects in the AddIns collection. Read-only Long.

## Methods (3)

- `Item(Index As Variant) As AddIn`  
  Returns an individual object in a collection.
    - `Index As Variant` (required): The individual add-in to be returned. Index can be an Long indicating the ordinal position of the add-in in the collection or a String representing the name of the individual add-in.
- `Add(FileName As String, [Install As Variant]) As AddIn`  
  Returns an AddIn object that represents an add-in added to the list of available add-ins.
    - `FileName As String` (required): The path for the template or WLL.
    - `Install As Variant` (optional): True to install the add-in. False to add the add-in to the list of add-ins but not install it. The default value is True.
- `Unload(RemoveFromList As Boolean)`  
  Unloads all loaded add-ins and, depending on the value of the RemoveFromList argument, removes them from the AddIns collection.
    - `RemoveFromList As Boolean` (required): True to remove the unloaded add-ins from the AddIns collection (the names are removed from the Templates and Add-ins dialog box). False to leave the unloaded add-ins in the collection. If the Autoload property for an unloaded add-in returns True, Unload cannot remove that add-in from the AddIns collection, regardless of the value of RemoveFromList.
