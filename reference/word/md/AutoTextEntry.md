# AutoTextEntry

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020936-0000-0000-C000-000000000046}  

Represents a single AutoText entry. The AutoTextEntry object is a member of the AutoTextEntries collection. The AutoTextEntries collection contains all the AutoText entries in the specified template. The entries are listed on the AutoText tab in the AutoCorrect dialog box.

**Remarks:** Use AutoTextEntries (_index_), where _index_ is the AutoText entry name or index number, to return a single AutoTextEntry object. You must exactly match the spelling (but not necessarily the capitalization) of the name, as it is shown on the AutoText tab in the AutoCorrect dialog box. The following example sets the value of an existing AutoText entry named "cName." The following example displays the name and value of the first AutoText entry in the template attached to the active document. The following example inserts the global AutoText entry named "TheWorld" at the insertion point. Use the Add method to add an AutoTextEntry object to the AutoTextEntries collection. The following example adds an AutoText entry named "Blue" based on the text of the selection.

## Properties (7)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified AutoTextEntries collection.
- `Index As Long  (read-only)`  
  Returns a Long that represents the position of an item in a collection. Read-only.
- `Name As String  (read/write)`  
  Returns or sets the name of the specified object.Read/write String.
- `StyleName As String  (read-only)`  
  Returns the name of the style applied to the specified AutoText entry. Read-only String.
- `Value As String  (read/write)`  
  Returns or sets the value of the AutoText entry. Read/write String.

## Methods (2)

- `Delete()`  
  Deletes the specified object.
- `Insert(Where As Range, [RichText As Variant]) As Range`  
  Inserts the AutoText entry in place of the specified range. Returns a Range object that represents the AutoText entry.
    - `Where As Range` (required): The location for the AutoText entry.
    - `RichText As Variant` (optional): True to insert the AutoText entry with its original formatting.
