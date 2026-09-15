# TwoInitialCapsException

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020943-0000-0000-C000-000000000046}  

Represents a single initial-capital AutoCorrect exception. The TwoInitialCapsException object is a member of the TwoInitialCapsExceptions collection. The TwoInitialCapsExceptions collection includes all the items listed in the Don't correct box on the INitial CAps tab in the AutoCorrect Exceptions dialog box.

**Remarks:** Use TwoInitialCapsExceptions (Index), where Index is the initial capital exception name or the index number, to return a single TwoInitialCapsException object. The following example deletes the initial-capital exception named "KMenu." The index number represents the position of the initial-capital exception in the TwoInitialCapsExceptions collection. The following example displays the name of the first item in the TwoInitialCapsExceptions collection. If the TwoInitialCapsAutoAdd property is True, words are automatically added to the list of initial-capital exceptions. Use the Add method to add an item to the TwoInitialCapsExceptions collection. The following example adds "INdustry" to the list of initial-capital exceptions.

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified TwoInitialCapsException object.
- `Index As Long  (read-only)`  
  Returns a Long that represents the position of an item in a collection. Read-only.
- `Name As String  (read-only)`  
  Returns the name of the specified object. Read-only String.

## Methods (1)

- `Delete()`  
  Deletes the specified two initial-capital exception from the collection of AutoCorrect exceptions.
