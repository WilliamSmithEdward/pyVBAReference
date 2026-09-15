# Dialogs

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020910-0000-0000-C000-000000000046}  

A collection of Dialog objects in Word. Each Dialog object represents a built-in Word dialog box.

**Remarks:** Use the Dialogs property to return the Dialogs collection. The following example displays the number of available built-in dialog boxes. You cannot create a new built-in dialog box or add one to the Dialogs collection. Use Dialogs (Index), where Index is the WdWordDialog constant that identifies the dialog box, to return a single Dialog object. The following example displays the built-in Open dialog box. For more information, see Displaying built-in Word dialog boxes.

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Dialogs object.
- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of dialog boxes in the collection. Read-only.

## Methods (1)

- `Item(Index As WdWordDialog) As Dialog`  
  Returns a dialog in Microsoft Word.
    - `Index As WdWordDialog` (required): A constant that specifies the dialog.
