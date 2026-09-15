# OMathAutoCorrectEntries

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {18CD5EC8-8B7B-42C8-992A-2A407468642C}  

Represents a collection of math AutoCorrect entries. Use the OMathAutoCorrectEntry object to access individual AutoCorrect entries.

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified OMathAutoCorrectEntries object.
- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns the number of items in the OMathAutoCorrectEntries collection. Read-only Long.

## Methods (2)

- `Item(Index As Variant) As OMathAutoCorrectEntry`  
  Returns an OMathAutoCorrectEntry object that represents the specified item in the collection.
    - `Index As Variant` (required): Specifies a String or Integer that represents the name or ordinal position of the object in the collection.
- `Add(Name As String, Value As String) As OMathAutoCorrectEntry`  
  Creates an equation auto correct entry and returns an OMathAutoCorrectEntry object.
    - `Name As String` (required): The name of the autocorrect entry. Corresponds to the Name property of the OMathAutoCorrectEntry object.
    - `Value As String` (required): The value of the autocorrect entry. Corresponds to the Value property of the OMathAutoCorrectEntry object.
