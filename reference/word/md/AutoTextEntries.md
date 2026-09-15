# AutoTextEntries

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020937-0000-0000-C000-000000000046}  

A collection of AutoCorrectEntry objects that represent the AutoText entries in a template. The AutoTextEntries collection includes all the entries listed on the AutoText tab in the AutoCorrect dialog box.

**Remarks:** Use the AutoTextEntries property to return the AutoTextEntries collection. The following example determines whether an AutoTextEntry object named "test" is in the AutoTextEntries collection. Use the Add method to add an AutoText entry to the AutoTextEntries collection. The following example adds an AutoText entry named "Blue" based on the text of the selection. Use AutoTextEntries (index), where index is the AutoText entry name or index number, to return a single AutoTextEntry object. You must exactly match the spelling (but not necessarily the capitalization) of the name, as it is shown on the AutoText tab in the AutoCorrect dialog box. The following example sets the value of an existing AutoText entry named "cName." The following example displays the name and value of the first AutoText entry in the template attached to the active document.

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified AutoTextEntries collection.
- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns the number of items in the AutoTextEntries collection. Read-only Long.

## Methods (3)

- `Item(Index As Variant) As AutoTextEntry`  
  Returns an individual AutoTextEntry object in a collection.
    - `Index As Variant` (required): The individual object to be returned. Can be a Long indicating the ordinal position or a String representing the name of the individual object.
- `Add(Name As String, Range As Range) As AutoTextEntry`  
  Returns an AutoTextEntry object that represents an AutoText entry added to the list of available AutoText entries.
    - `Name As String` (required): The text that, when typed, initiates an AutoText entry.
    - `Range As Range` (required): A range of text that will be inserted whenever Name is typed.
- `AppendToSpike(Range As Range) As AutoTextEntry`  
  Deletes the specified range and adds the contents of the range to the Spike (a built-in AutoText entry). This method returns the Spike as an AutoTextEntry object.
    - `Range As Range` (required): The range that's deleted and appended to the Spike.
