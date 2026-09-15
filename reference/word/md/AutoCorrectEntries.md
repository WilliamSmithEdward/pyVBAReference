# AutoCorrectEntries

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020948-0000-0000-C000-000000000046}  

A collection of AutoCorrectEntry objects that represent all the AutoCorrect entries available to Word. The AutoCorrectEntries collection includes all the entries in the AutoCorrect dialog box.

**Remarks:** Use the Entries property to return the AutoCorrectEntries collection. The following example displays the number of AutoCorrectEntry objects in the AutoCorrectEntries collection. Use the Add or AddRichText method to add an AutoCorrect entry to the list of available entries. The following example adds a plain-text AutoCorrect entry for the misspelling of the word "their." The following example creates an AutoCorrect entry named "PMO" based on the text and formatting of the selection. Use Entries (index), where index is the AutoCorrect entry name or index number, to return a single AutoCorrectEntry object. You must exactly match the spelling (but not necessarily the capitalization) of the name, as it is shown under Replace in the AutoCorrect dialog box. The following example sets the value of an existing AutoCorrect entry named "teh." The following example displays the name and value of the first AutoCorrent entry.

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified AutoCorrectEntries collection, which is usually an AutoCorrect object.
- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns the number of items in the AutoCorrectEntries collection. Read-only Long.

## Methods (3)

- `Item(Index As Variant) As AutoCorrectEntry`  
  Returns an individual AutoCorrectEntry object in a collection.
    - `Index As Variant` (required): The individual object to be returned. Can be a Long indicating the ordinal position or a String representing the name of the individual object.
- `Add(Name As String, Value As String) As AutoCorrectEntry`  
  Returns an AutoCorrectEntry object that represents a plain-text AutoCorrect entry added to the list of available AutoCorrect entries.
    - `Name As String` (required): The text you want to have automatically replaced with the text specified by Value.
    - `Value As String` (required): The text you want to have automatically inserted whenever the text specified by Name is typed.
- `AddRichText(Name As String, Range As Range) As AutoCorrectEntry`  
  Creates a formatted AutoCorrect entry, preserving all text attributes of the specified range. Returns an AutoCorrectEntry object.
    - `Name As String` (required): The text to replace automatically with Range.
    - `Range As Range` (required): The formatted text that Word will insert automatically whenever Name is typed.
