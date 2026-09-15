# OtherCorrectionsExceptions

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {000209DF-0000-0000-C000-000000000046}  

A collection of OtherCorrectionsException objects that represents the list of words that Microsoft Word won't correct automatically.

**Remarks:** This list corresponds to the list of AutoCorrect exceptions on the Other Corrections tab in the AutoCorrect Exceptions dialog box. Use the OtherCorrectionsExceptions property to return the OtherCorrectionsExceptions collection. The following example displays the items in this collection. If the value of the OtherCorrectionsAutoAdd property is True, words are automatically added to the list of AutoCorrect exceptions. Use the Add method to add an item to the OtherCorrectionsExceptions collection. The following example adds "TipTop" to the list of AutoCorrect exceptions. Use OtherCorrectionsExceptions (Index), where Index is the name or the index number, to return a single OtherCorrectionsException object. The following example deletes "WTop" from the list of AutoCorrect exceptions. The index number represents the position of the AutoCorrect exception in the OtherCorrectionsExceptions collection. The following example displays the name of the first item in the OtherCorrectionsExceptions collection.

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified OtherCorrectionsExceptions object.
- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of exceptions in the collection. Read-only.

## Methods (2)

- `Item(Index As Variant) As OtherCorrectionsException`  
  Returns an individual OtherCorrectionsException object in a collection.
    - `Index As Variant` (required): The individual object to be returned. Can be a Long indicating the ordinal position or a String representing the name of the individual object.
- `Add(Name As String) As OtherCorrectionsException`  
  Returns an OtherCorrectionsException object that represents a new exception added to the list of AutoCorrect exceptions.
    - `Name As String` (required): The word that you want Word to overlook.
