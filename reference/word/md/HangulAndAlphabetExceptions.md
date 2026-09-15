# HangulAndAlphabetExceptions

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {000209D1-0000-0000-C000-000000000046}  

A collection of HangulAndAlphabetException objects that represents all Hangul and alphabet AutoCorrect exceptions.

**Remarks:** Use the HangulAndAlphabetExceptions property to return the HangulAndAlphabetExceptions collection. The following example displays the items in this collection. If the value of the HangulAndAlphabetAutoAdd property is True, words are automatically added to the list of Hangul and alphabet AutoCorrect exceptions. Use the Add method to add an item to the HangulAndAlphabetExceptions collection. The following example adds "hello" to the list of alphabet AutoCorrect exceptions. Use HangulAndAlphabetExceptions (Index), where Index is the Hangul or alphabet AutoCorrect exception name or the index number, to return a single HangulAndAlphabetException object. The following example deletes the alphabet AutoCorrect exception named "goodbye." The index number represents the position of the hangul or alphabet AutoCorrect exception in the HangulAndAlphabetExceptions collection. The following example displays the name of the first item in the HangulAndAlphabetExceptions collection.

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified HangulAndAlphabetExceptions object.
- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of exceptions in the collection. Read-only.

## Methods (2)

- `Item(Index As Variant) As HangulAndAlphabetException`  
  Returns an individual HangulAndAlphabetException object in a collection.
    - `Index As Variant` (required): The individual object to be returned. Can be a Long indicating the ordinal position or a String representing the name of the individual object.
- `Add(Name As String) As HangulAndAlphabetException`  
  Returns a HangulAndAlphabetException object that represents a new exception to the list of AutoCorrect exceptions.
    - `Name As String` (required): The word that you don't want Microsoft Word to correct automatically.
