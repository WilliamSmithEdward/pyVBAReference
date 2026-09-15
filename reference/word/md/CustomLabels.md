# CustomLabels

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020916-0000-0000-C000-000000000046}  

A collection of CustomLabel objects available in the Label Options dialog box. This collection includes custom labels of all printer types (dot-matrix, laser, and ink-jet printers).

**Remarks:** Use the CustomLabels property to return the CustomLabels collection. The following example displays the number of available custom labels. Use the Add method to create a custom label. The following example adds a custom mailing label named "My Label" and sets the page size. Use CustomLabels (Index), where Index is the custom label name or index number, to return a single CustomLabel object. The following example creates a new document with an existing custom label layout named "My Labels." The index number represents the position of the custom mailing label in the CustomLabels collection. The following example displays the name of the first custom mailing label.

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified CustomLabels object.
- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of items in the collection. Read-only.

## Methods (2)

- `Item(Index As Variant) As CustomLabel`  
  Returns a CustomLabel object in a collection.
    - `Index As Variant` (required): The individual object to be returned. Can be a Long indicating the ordinal position or a String representing the name of the individual object.
- `Add(Name As String, [DotMatrix As Variant]) As CustomLabel`  
  Adds a custom mailing label to the CustomLabels collection. Returns a CustomLabel object that represents the custom mailing label.
    - `Name As String` (required): The name for the custom mailing labels.
    - `DotMatrix As Variant` (optional): True to have the mailing labels printed on a dot-matrix printer.
