# CaptionLabels

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {00020978-0000-0000-C000-000000000046}  

A collection of CaptionLabel objects that represent the available caption labels. The items in the CaptionLabels collection are listed in the Label box in the Caption dialog box.

**Remarks:** Use the CaptionLabels property to return the CaptionLabels collection. By default, the CaptionLabels collection includes the three built-in caption labels: Figure, Table, and Equation. Use the Add method to add a custom caption label. The following example adds a caption label named "Photo." Use CaptionLabels (_index_), where _index_ is the caption label name or index number, to return a single CaptionLabel object. The following example sets the numbering style for the Figure caption label. The index number represents the position of the caption label in the CaptionLabels collection. The following example displays the first caption label.

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified CaptionLabels object.
- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of caption labels in the collection. Read-only.

## Methods (2)

- `Item(Index As Variant) As CaptionLabel`  
  Returns an individual CaptionLabel object in a collection.
    - `Index As Variant` (required): The individual object to be returned. Can be a Long indicating the ordinal position or a String representing the name of the individual object.
- `Add(Name As String) As CaptionLabel`  
  Returns a CaptionLabel object that represents a custom caption label.
    - `Name As String` (required): The name of the custom caption label.
