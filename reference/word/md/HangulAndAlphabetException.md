# HangulAndAlphabetException

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {000209D2-0000-0000-C000-000000000046}  

Represents a single Hangul or alphabet AutoCorrect exception. The HangulAndAlphabetException object is a member of the HangulAndAlphabetExceptions collection.

**Remarks:** Use HangulAndAlphabetExceptions (Index), where Index is the Hangul or alphabet AutoCorrect exception name or the index number, to return a single HangulAndAlphabetException object. The following example deletes the alphabet AutoCorrect exception named "hello." The index number represents the position of the Hangul or alphabet AutoCorrect exception in the HangulAndAlphabetExceptions collection. The following example displays the name of the first item in the HangulAndAlphabetExceptions collection. If the value of the HangulAndAlphabetAutoAdd property is True, words are automatically added to the list of Hangul and alphabet AutoCorrect exceptions. Use the Add method to add an item to the HangulAndAlphabetExceptions collection. The following example adds "goodbye" to the list of alphabet AutoCorrect exceptions.

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified HangulAndAlphabetException object.
- `Index As Long  (read-only)`  
  Returns a Long that represents the position of an item in a collection. Read-only.
- `Name As String  (read-only)`  
  Returns the name of the specified object. Read-only String.

## Methods (1)

- `Delete()`  
  Deletes the specified exception.
