# CoAuthLock

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {99755F80-FE96-4F7D-B636-B8E800E54F44}  

Represents a lock within the document. The CoAuthLock object is a member of the CoAuthLocks collection.

**Remarks:** Use Locks (_index_), where _index_ is the index number, to return a CoAuthLock object. When adding a CoAuthLock object, use the WdLockType enumeration to specify the type of lock.

**Example:**

```vba
Dim myLock as CoAuthLock

Set myLock = ActiveDocument.CoAuthoring.Locks(1)
```

## Properties (7)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application. Read-only.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified CoAuthLock object.
- `Type As WdLockType  (read-only)`  
  Returns a WdLockType constant that specifies the lock type. Read-only.
- `Owner As CoAuthor  (read-only)`
- `Range As Range  (read-only)`  
  Returns a Range object that represents the portion of a document that is contained in the specified object. Read-only.
- `HeaderFooter As Boolean  (read-only)`

## Methods (1)

- `Unlock()`  
  Unlocks the specified lock.
