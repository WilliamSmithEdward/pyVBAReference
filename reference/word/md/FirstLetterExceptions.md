# FirstLetterExceptions

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020946-0000-0000-C000-000000000046}  

A collection of FirstLetterException objects that represent the abbreviations excluded from automatic correction.

**Remarks:** The first character following a period is automatically capitalized when the CorrectSentenceCaps property is set to True. The FirstLetterExceptions collection includes exceptions to this behavior (for example, abbreviations such as "addr." and "apt."). Use the FirstLetterExceptions property to return the FirstLetterExceptions collection. The following example deletes the abbreviation "addr." if it is included in the FirstLetterExceptions collection. The following example creates a new document and inserts all the AutoCorrect first-letter exceptions into it. Use the Add method to add an abbreviation to the list of first-letter exceptions. The following example adds the abbreviation "addr." to this list. Use FirstLetterExceptions (Index), where Index is the abbreviation or the index number, to return a single FirstLetterException object. The following example deletes the abbreviation "appt." from the FirstLetterExceptions collection. The following example displays the name of the first item in the FirstLetterExceptions collection.

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified FirstLetterExceptions object.
- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of first letter exceptions in the collection. Read-only.

## Methods (2)

- `Item(Index As Variant) As FirstLetterException`  
  Returns an individual FirstLetterException object in a collection.
    - `Index As Variant` (required): The individual object to be returned. Can be a Long indicating the ordinal position or a String representing the name of the individual object.
- `Add(Name As String) As FirstLetterException`  
  Returns a FirstLetterException object that represents a new exception added to the list of AutoCorrect exceptions.
    - `Name As String` (required): The word with two initial capital letters that you want Microsoft Word to overlook.
