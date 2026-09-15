# Dialog

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {000209B8-0000-0000-C000-000000000046}  

Represents a built-in dialog box. The Dialog object is a member of the Dialogs collection. The Dialogs collection contains all the built-in dialog boxes in Word. You cannot create a new built-in dialog box or add one to the Dialogs collection.

**Remarks:** Use Dialogs (Index), where Index is a WdWordDialog constant that identifies the dialog box, to return a single Dialog object. The following example displays and carries out the actions taken in the built-in Open dialog box. The WdWordDialog constants are formed from the prefix "wdDialog" followed by the name of the menu and the dialog box. For example, the constant for the Page Setup dialog box is wdDialogFilePageSetup, and the constant for the New dialog box is wdDialogFileNew. For more information about working with built-in Word dialog boxes, see Displaying built-in Word dialog boxes.

## Properties (7)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Dialog object.
- `DefaultTab As WdWordDialogTab  (read/write)`  
  Returns or sets the active tab when the specified dialog box is displayed. Read/write WdWordDialogTab.
- `Type As WdWordDialog  (read-only)`  
  Returns the type of built-in Microsoft Word dialog box. Read-only WdWordDialog.
- `CommandName As String  (read-only)`  
  Returns the name of the procedure that displays the specified built-in dialog box. Read-only String.
- `CommandBarId As Long  (read-only)`  
  Returns a Long that represents the toolbar control id for a built-in Microsoft Word dialog box. Read-only.

## Methods (4)

- `Show([TimeOut As Variant]) As Long`  
  Displays and carries out actions initiated in the specified built-in Word dialog box. Returns a Long that indicates which button was clicked to close the dialog box.
    - `TimeOut As Variant` (optional): The amount of time that Word will wait before closing the dialog box automatically. One unit is approximately 0.001 second. Concurrent system activity may increase the effective time value. If this argument is omitted, the dialog box is closed when the user dismisses it.
- `Display([TimeOut As Variant]) As Long`  
  Displays the specified built-in Word dialog box until either the user closes it or the specified amount of time has passed. Returns a Long that indicates which button was clicked to close the dialog box.
    - `TimeOut As Variant` (optional): The amount of time that Word will wait before closing the dialog box automatically. One unit is approximately 0.001 second. Concurrent system activity may increase the effective time value. If this argument is omitted, the dialog box is closed when the user closes it.
- `Execute()`  
  Applies the current settings of a Microsoft Word dialog box.
- `Update()`  
  Updates the values shown in a built-in Microsoft Word dialog box.
