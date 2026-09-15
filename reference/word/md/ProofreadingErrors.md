# ProofreadingErrors

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {000209BB-0000-0000-C000-000000000046}  

A collection of spelling and grammatical errors for the specified document or range.

**Remarks:** Use the SpellingErrors or GrammaticalErrors property to return the ProofreadingErrors collection. The following example counts the spelling and grammatical errors in the selection and displays the results in a message box. Use SpellingErrors (Index), where Index is the index number, to return a single spelling error (represented by a Range object). The following example finds the second spelling error in the selection and then selects it. Use GrammarErrors (Index), where Index is the index number, to return a single grammatical error (represented by a Range object). The following example returns the sentence that contains the first grammatical error in the selection. The Count property for this collection in a document returns the number of items in the main story only. To count items in other stories use the collection with the Range object. If all the words in the document or range are spelled correctly and are grammatically correct, the Count property for the ProofreadingErrors object returns 0 (zero) and the SpellingChecked and GrammarChecked properties return True.

## Properties (6)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified ProofreadingErrors object.
- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of proofreading errors in the collection. Read-only.
- `Type As WdProofreadingErrorType  (read-only)`  
  Returns the type of proofreading error. Read-only WdProofreadingErrorType.

## Methods (1)

- `Item(Index As Long) As Range`  
  Returns an individual Range object in a collection.
    - `Index As Long` (required): The individual object to be returned. Can be a Long indicating the ordinal position of the individual object.
