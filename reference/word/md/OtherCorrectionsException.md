# OtherCorrectionsException

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {000209E1-0000-0000-C000-000000000046}  

Represents a single AutoCorrect exception. The OtherCorrectionsException object is a member of the OtherCorrectionsExceptions collection.

**Remarks:** The OtherCorrectionsExceptions collection includes all words that Microsoft Word won't correct automatically. This list corresponds to the list of AutoCorrect exceptions on the Other Corrections tab in the AutoCorrect Exceptions dialog box. Use OtherCorrectionsExceptions (Index), where Index is the AutoCorrect exception name or the index number, to return a single OtherCorrectionsException object. The following example deletes "WTop" from the list of AutoCorrect exceptions. The index number represents the position of the AutoCorrect exception in the OtherCorrectionsExceptions collection. The following example displays the name of the first item in the OtherCorrectionsExceptions collection. If the value of the OtherCorrectionsAutoAdd property is True, words are automatically added to the list of AutoCorrect exceptions. Use the Add method to add an item to the OtherCorrectionsExceptions collection. The following example adds "TipTop" to the list of AutoCorrect exceptions.

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified OtherCorrectionsException object.
- `Index As Long  (read-only)`  
  Returns a Long that represents the position of an item in a collection. Read-only.
- `Name As String  (read-only)`  
  Returns the name of the specified object. Read-only String.

## Methods (1)

- `Delete()`  
  Deletes the specified exception.
