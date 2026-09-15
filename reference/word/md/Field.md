# Field

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {0002092F-0000-0000-C000-000000000046}  

Represents a field. The Field object is a member of the Fields collection. The Fields collection represents the fields in a selection, range, or document.

**Remarks:** Use Fields (Index), where Index is the index number, to return a single Field object. The index number represents the position of the field in the selection, range, or document. The following example displays the field code and the result of the first field in the active document. Use the Add method to add a field to the Fields collection. The following example inserts a DATE field at the beginning of the selection and then displays the result.

## Properties (16)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Field object.
- `Code As Range  (read/write)`  
  Returns a Range object that represents a field's code. Read/write.
- `Type As WdFieldType  (read-only)`  
  Returns the field type. Read-only WdFieldType.
- `Locked As Boolean  (read/write)`  
  True if the specified field is locked. Read/write Boolean.
- `Kind As WdFieldKind  (read-only)`  
  Returns the type of link for a Field object. Read-only WdFieldKind.
- `Result As Range  (read/write)`  
  Returns a Range object that represents a field's result. Read/write.
- `Data As String  (read/write)`  
  Returns or sets data in an ADDIN field. Read/write String.
- `Next As Field  (read-only)`  
  Returns the next object in the collection. Read-only.
- `Previous As Field  (read-only)`  
  Returns the previous object in the collection. Read-only.
- `Index As Long  (read-only)`  
  Returns a Long that represents the position of an item in a collection. Read-only.
- `ShowCodes As Boolean  (read/write)`  
  True if field codes are displayed for the specified field instead of field results. Read/write Boolean.
- `LinkFormat As LinkFormat  (read-only)`  
  Returns a LinkFormat object that represents the link options of the specified field. Read/only.
- `OLEFormat As OLEFormat  (read-only)`  
  Returns an OLEFormat object that represents the OLE characteristics (other than linking) for the specified field. Read-only.
- `InlineShape As InlineShape  (read-only)`  
  Returns an InlineShape object that represents the picture, OLE object, or ActiveX control that is the result of an INCLUDEPICTURE or EMBED field.

## Methods (8)

- `Select()`  
  Selects the specified field.
- `Update() As Boolean`  
  Updates the result of the field. Returns True if the field is updated successfully.
- `Unlink()`  
  Replaces the specified field with its most recent result.
- `UpdateSource()`  
  Saves the changes made to the results of an INCLUDETEXT field back to the source document.
- `DoClick()`  
  Clicks the specified field.
- `Copy()`  
  Copies the specified field to the Clipboard.
- `Cut()`  
  Removes the specified field from the document and places it on the Clipboard.
- `Delete()`  
  Deletes the specified field.
