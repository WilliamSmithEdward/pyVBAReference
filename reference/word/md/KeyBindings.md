# KeyBindings

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020996-0000-0000-C000-000000000046}  

A collection of KeyBinding objects that represent the custom key assignments in the current context. Custom key assignments are made in the Customize Keyboard dialog box.

**Remarks:** Use the KeyBindings property to return the KeyBindings collection. The following example inserts after the selection the command name and key combination for each item in the KeyBindings collection. Use the Add method to add a KeyBinding object to the KeyBindings collection. The following example adds the CTRL+ALT+H key combination to the Heading 1 style in the active document. Use KeyBindings (Index), where Index is the index number, to return a single KeyBinding object. The following example displays the command associated with the first KeyBinding object in the KeyBindings collection.

## Properties (6)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified KeyBindings object.
- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of key bindings in the collection. Read-only.
- `Context As Object  (read-only)`  
  Returns an Object that represents the storage location of the specified key binding. Read-only.

## Methods (4)

- `Item(Index As Long) As KeyBinding`  
  Returns an individual KeyBinding object in a collection.
    - `Index As Long` (required): The individual object to be returned. Can be a Long indicating the ordinal position of the individual object.
- `Add(KeyCategory As WdKeyCategory, Command As String, KeyCode As Long, [KeyCode2 As Variant], [CommandParameter As Variant]) As KeyBinding`  
  Returns a KeyBinding object that represents a new shortcut key for a macro, built-in command, font, AutoText entry, style, or symbol.
    - `KeyCategory As WdKeyCategory` (required): The category of the key assignment.
    - `Command As String` (required): The command that the specified key combination executes.
    - `KeyCode As Long` (required): A key you specify by using one of the WdKey constants.
    - `KeyCode2 As Variant` (optional): A second key you specify by using one of the WdKey constants.
    - `CommandParameter As Variant` (optional): Additional text, if any, required for the command specified by Command. For details, see the Remarks section below.
- `ClearAll()`  
  Clears all the customized key assignments and restores the original Microsoft Word shortcut key assignments.
- `Key(KeyCode As Long, [KeyCode2 As Variant]) As KeyBinding`  
  Returns a KeyBinding object that represents the specified custom key combination.
    - `KeyCode As Long` (required): A key you specify by using one of the WdKey constants.
    - `KeyCode2 As Variant` (optional): A second key you specify by using one of the WdKey constants.
