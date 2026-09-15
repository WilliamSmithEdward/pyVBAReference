# TwoInitialCapsExceptions

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020944-0000-0000-C000-000000000046}  

A collection of TwoInitialCapsException objects that represent all the items listed in the Don't correct box on the INitial CAps tab in the AutoCorrect Exceptions dialog box.

**Remarks:** Use the TwoInitialCapsExceptions property to return the TwoInitialCapsExceptions collection. The following example displays the items in this collection. If the TwoInitialCapsAutoAdd property is True, words are automatically added to the list of initial-capital exceptions. Use the Add method to add an item to the TwoInitialCapsExceptions collection. The following example adds "Industry" to the list of initial-capital exceptions. Use TwoInitialCapsExceptions (Index), where Index is the initial cap name or the index number, to return a single TwoInitialCapsException object. The following example deletes the initial-capital item named "KMenu." The index number represents the position of the initial-capital exception in the TwoInitialCapsExceptions collection. The following example displays the name of the first item in the TwoInitialCapsExceptions collection.

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified TwoInitialCapsExceptions object.
- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of exceptions in the collection. Read-only.

## Methods (2)

- `Item(Index As Variant) As TwoInitialCapsException`  
  Returns an individual TwoInitialCapsException object in a collection.
    - `Index As Variant` (required): The individual object to be returned. Can be a Long indicating the ordinal position or a String representing the name of the individual object.
- `Add(Name As String) As TwoInitialCapsException`  
  Returns a TwoInitialCapsException object that represents a new exception added to the list of AutoCorrect exceptions.
    - `Name As String` (required): The word with two initial capital letters that you want Microsoft Word to overlook.
