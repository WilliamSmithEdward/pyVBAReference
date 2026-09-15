# AddIn

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {0002097E-0000-0000-C000-000000000046}  

Represents a single add-in, either installed or not installed. The AddIn object is a member of the AddIns collection. The AddIns collection contains all the add-ins available to Microsoft Word, regardless of whether they are currently loaded. The AddIns collection includes global templates or Word add-in libraries (WLLs) displayed in the Templates and Add-ins dialog box.

**Remarks:** Use AddIns (index), where index is the add-in name or index number, to return a single AddIn object. You must exactly match the spelling (but not necessarily the capitalization) of the name, as it is shown in the Templates and Add-Ins dialog box. The following example loads the Letter.dot template as a global template. The index number represents the position of the add-in in the list of add-ins in the Templates and Add-ins dialog box. The following instruction displays the path of the first available add-in. The following example creates a list of add-ins at the beginning of the active document. The list contains the name, path, and installed state of each available add-in. Use the Add method to add an add-in to the list of available add-ins and (optionally) install it using the Install argument. To install an add-in shown in the list of available add-ins, use the Installed property.

## Properties (9)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified AddIn object.
- `Name As String  (read-only)`  
  Returns the name of an add-in. Read-only String.
- `Index As Long  (read-only)`  
  Returns a Long that represents the position of an item in a collection. Read-only.
- `Path As String  (read-only)`  
  Returns the location of an installed add-in. Read-only String.
- `Installed As Boolean  (read/write)`  
  True if the specified add-in is installed (loaded). Add-ins that are loaded are selected in the Templates and Add-ins dialog box. Read/write Boolean.
- `Compiled As Boolean  (read-only)`  
  True if the specified add-in is a Word add-in library (WLL). False if the add-in is a template. Read-only Boolean.
- `Autoload As Boolean  (read-only)`  
  True if the specified add-in is automatically loaded when Word is started. Add-ins located in the Startup folder in the Word program folder are automatically loaded. Read-only Boolean.

## Methods (1)

- `Delete()`  
  Deletes the specified add-in.
