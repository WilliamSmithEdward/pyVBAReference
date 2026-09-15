# AutoCorrectEntry

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020947-0000-0000-C000-000000000046}  

Represents a single AutoCorrect entry. The AutoCorrectEntry object is a member of the AutoCorrectEntries collection. The AutoCorrectEntries collection includes the entries in the AutoCorrect dialog box.

**Remarks:** Use Entries (index), where index is the AutoCorrect entry name or index number, to return a single AutoCorrectEntry object. You must exactly match the spelling (but not necessarily the capitalization) of the name, as it is shown under Replace in the AutoCorrect dialog box. The following example sets the value of the AutoCorrect entry named "teh." Use the Apply method to insert an AutoCorrect entry at the specified range. The following example adds an AutoCorrect entry and then inserts it in place of the selection. Use either the Add or AddRichText method to add an AutoCorrect entry to the list of available entries. The following example adds a plain-text AutoCorrect entry for the misspelling of the word "their.' The following example creates an AutoCorrect entry named "PMO" based on the text and formatting of the selection.

## Properties (7)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified AutoCorrectEntry object.
- `Index As Long  (read-only)`  
  Returns a Long that represents the position of an item in a collection. Read-only.
- `Name As String  (read/write)`  
  Returns or sets the name of the specified object. Read/write String.
- `Value As String  (read/write)`  
  Returns or sets the value of the AutoCorrect entry. Read/write String.
- `RichText As Boolean  (read-only)`  
  True if formatting is stored with the AutoCorrect entry replacement text. Read-only Boolean.

## Methods (2)

- `Delete()`  
  Deletes the specified AutoCorrect entry.
- `Apply(Range As Range)`  
  Replaces a range with the value of the specified AutoCorrect entry.
    - `Range As Range` (required): The range to which to apply the AutoCorrect entry.
