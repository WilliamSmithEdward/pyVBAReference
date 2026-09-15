# KeysBoundTo

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020997-0000-0000-C000-000000000046}  

A collection of KeyBinding objects assigned to a command, style, macro, or other item in the current context.

**Remarks:** Use the KeysBoundTo property to return the KeysBoundTo collection. The following example displays the key combinations assigned to the FileNew command in the Normal template. The following example displays the name of the document or template where the keys for the macro named "Macro1" are stored.

## Properties (9)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified KeysBoundTo object.
- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of keys in the collection. Read-only.
- `KeyCategory As WdKeyCategory  (read-only)`  
  Returns the type of item assigned to the specified key binding. Read-only WdKeyCategory.
- `Command As String  (read-only)`  
  Returns a String that represents the command assigned to the specified key combination. Read-only.
- `CommandParameter As String  (read-only)`  
  Returns the command parameter assigned to the specified shortcut key. Read-only String.
- `Context As Object  (read-only)`  
  Returns an Object that represents the storage location of the specified key binding. Read-only.

## Methods (2)

- `Item(Index As Long) As KeyBinding`  
  Returns an individual KeyBinding object in a collection.
    - `Index As Long` (required): The individual object to be returned. Can be a Long indicating the ordinal position of the individual object.
- `Key(KeyCode As Long, [KeyCode2 As Variant]) As KeyBinding`  
  Returns a KeyBinding object that represents the specified custom key combination.
    - `KeyCode As Long` (required): A key you specify by using one of the WdKey constants.
    - `KeyCode2 As Variant` (optional): A second key you specify by using one of the WdKey constants.
