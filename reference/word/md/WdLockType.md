# WdLockType

**Type:** Enumeration  
**Library:** Microsoft Word 16.0 Object Library  

Specifies the type of lock for a CoAuthLock object.

## Constants (4)

- `wdLockNone` = 0  
  Reserved for future use.
- `wdLockReservation` = 1  
  Specifies a reservation lock. A reservation lock is explicitly created by a user through the Block Authors button on the Review tab in Word.
- `wdLockEphemeral` = 2  
  Specifies an ephemeral lock. Word implicitly places an ephemeral lock on a range when a user begins editing a range in a document with coauthoring enabled.
- `wdLockChanged` = 3  
  Specifies a placeholder lock. A placeholder lock indicates that another user has removed their lock from the range, but the current user has not updated their view of the document by saving.
