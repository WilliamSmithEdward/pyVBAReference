# FirstLetterException

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020945-0000-0000-C000-000000000046}  

Represents an abbreviation excluded from automatic correction. The FirstLetterExceptions object is a member of the FirstLetterExceptions collection.

**Remarks:** The FirstLetterExceptions collection includes all the excluded abbreviations.The first character following a period is automatically capitalized when the CorrectSentenceCaps property is set to True. The character you type following an item in the FirstLetterExceptions collection isn't capitalized. Use FirstLetterExceptions (Index), where Index is the abbreviation or the index number, to return a single FirstLetterException object. The following example deletes the abbreviation "appt." from the FirstLetterExceptions collection. The following example displays the name of the first item in the FirstLetterExceptions collection. Use the Add method to add an abbreviation to the list of first-letter exceptions. The following example adds the abbreviation "addr." to this list.

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified FirstLetterException object.
- `Index As Long  (read-only)`  
  Returns a Long that represents the position of an item in a collection. Read-only.
- `Name As String  (read-only)`  
  Returns name of the specified object. Read-only String.

## Methods (1)

- `Delete()`  
  Deletes the FirstLetterException object.
