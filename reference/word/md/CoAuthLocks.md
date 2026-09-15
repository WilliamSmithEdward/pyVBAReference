# CoAuthLocks

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {DFF99AC2-CD2A-43AD-91B1-A2BE40BC7146}  

A collection of CoAuthLock objects.

**Remarks:** Use the Locks property to return the CoAuthLocks collection.

**Example:**

```vba
MsgBox ActiveDocument.CoAuthoring.Locks.Count
```

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application. Read-only.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified CoAuthLocks object.
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of locks in the CoAuthLocks collection. Read-only.
- `_NewEnum As IUnknown  (read-only)`

## Methods (3)

- `Item(Index As Long) As CoAuthLock`  
  Returns an individual CoAuthLock object in a collection.
    - `Index As Long` (required): The individual object to be returned.
- `Add([Range As Variant], [Type As WdLockType]) As CoAuthLock`  
  Returns a CoAuthLock object that represents a lock added to a specified range.
    - `Range As Variant` (optional): Specifies the document range locked by the CoAuthLock object. This parameter may be a Paragraph, Column, Cell, Row, Table, Range, or Selection object.
    - `Type As WdLockType` (optional): Specifies the type of lock. The WdLockType specified can only be wdLockEphemeral or WdLockReservation
- `RemoveEphemeralLocks()`  
  Removes ephemeral locks from the document.
