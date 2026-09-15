# Fields

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020930-0000-0000-C000-000000000046}  

A collection of Field objects that represent all the fields in a selection, range, or document.

**Remarks:** Use the Fields property to return the Fields collection. The following example updates all the fields in the selection. Use the Add method to add a field to the Fields collection. The following example inserts a DATE field at the beginning of the selection and then displays the result. Use Fields (Index), where Index is the index number, to return a single Field object. The index number represents the position of the field in the selection, range, or document. The following example displays the field code and the result of the first field in the active document. The Count property for this collection in a document returns the number of items in the main story only. To count items in other stories use the collection with the Range object.

## Properties (6)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Fields object.
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of fields in the collection. Read-only.
- `Locked As Long  (read/write)`  
  True if all fields in the Fields collection are locked. Read/write Long.
- `_NewEnum As IUnknown  (read-only)`

## Methods (6)

- `Item(Index As Long) As Field`  
  Returns an individual Field object in a collection.
    - `Index As Long` (required): The individual object to be returned. Can be a Long indicating the ordinal position of the individual object.
- `ToggleShowCodes()`  
  Switches the display of the fields between field codes and field results. Use the ShowCodes property to control the display of an individual field.
- `Update() As Long`  
  Updates the result of the fields object.
- `Unlink()`  
  Replaces all the fields in the Fields collection with their most recent results.
- `UpdateSource()`  
  Saves the changes made to the results of an INCLUDETEXT field back to the source document.
- `Add(Range As Range, [Type As Variant], [Text As Variant], [PreserveFormatting As Variant]) As Field`  
  Adds a Field object to the Fields collection. Returns the Field object at the specified range.
    - `Range As Range` (required): The range where you want to add the field. If the range isn't collapsed, the field replaces the range.
    - `Type As Variant` (optional): Can be any WdFieldType constant. For a list of valid constants, consult the Object Browser. The default value is wdFieldEmpty.
    - `Text As Variant` (optional): Additional text needed for the field. For example, if you want to specify a switch for the field, you would add it here.
    - `PreserveFormatting As Variant` (optional): True to have the formatting that's applied to the field preserved during updates.
