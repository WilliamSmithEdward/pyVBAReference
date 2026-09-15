# KeyBinding

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020998-0000-0000-C000-000000000046}  

Represents a custom key assignment in the current context. The KeyBinding object is a member of the KeyBindings collection.

**Remarks:** Use KeyBindings (Index), where Index is the index number, to return a single KeyBinding object. The following example displays the command associated with the first KeyBinding object in the KeyBindings collection. You can also use the FindKey property and the Key method to return a KeyBinding object.

## Properties (11)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified KeyBinding object.
- `Command As String  (read-only)`  
  Returns the command assigned to the specified key combination. Read-only String.
- `KeyString As String  (read-only)`  
  Returns the key combination string for the specified keys (for example, CTRL+SHIFT+A). Read-only String.
- `Protected As Boolean  (read-only)`  
  True if you cannot change the specified key binding in the Customize Keyboard dialog box. Read-only Boolean.
- `KeyCategory As WdKeyCategory  (read-only)`  
  Returns the type of item assigned to the specified key binding. Read-only WdKeyCategory.
- `KeyCode As Long  (read-only)`  
  Returns a unique number for the first key in the specified key binding. Read-only Long.
- `KeyCode2 As Long  (read-only)`  
  Returns a unique number for the second key in the specified key binding. Read-only Long.
- `CommandParameter As String  (read-only)`  
  Returns the command parameter assigned to the specified shortcut key. Read-only String.
- `Context As Object  (read-only)`  
  Returns an Object that represents the storage location of the specified key binding. Read-only.

## Methods (4)

- `Clear()`  
  Removes the specified key binding from the KeyBindings collection and resets a built-in command to its default key assignment.
- `Disable()`  
  Removes the specified key combination if it is currently assigned to a command. After you use this method, the key combination has no effect.
- `Execute()`  
  Runs the command associated with the specified key combination.
- `Rebind(KeyCategory As WdKeyCategory, Command As String, [CommandParameter As Variant])`  
  Changes the command assigned to the specified key binding.
    - `KeyCategory As WdKeyCategory` (required): The key category of the specified key binding.
    - `Command As String` (required): The name of the specified command.
    - `CommandParameter As Variant` (optional): Additional text, if any, required for the command specified by Command. For information about values for this argument, see the Add method.
