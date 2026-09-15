# ContentControlListEntry

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {0C6FA8CA-E65F-4FC7-AB8F-20729EECBB14}  

A ContentControlListEntry object represents a list item in a drop-down list or combo box content control. A ContentControlListEntry object is a member of the ContentControlListEntries collection for a ContentControl object.

**Remarks:** Use the Add method of the ContentControlListEntries collection to create a new ContentControlListEntry object. Use the Item method, or DropdownListEntries (Index), where Index is the ordinal position of the content control list item, to access an individual list item within the ContentControlListEntries collection. The following code example uses the Add method to add several list items to a new drop-down list content control, and then uses the Item method to access the third item in the list and change the display text. Use the MoveUp and MoveDown methods to reposition items in a drop-down list. The following code example moves the first item down, so that it becomes the last item in the list, and moves the last item up, so that it becomes the first item in the list. Use the Select method to programmatically select a content control list item. The following code example inserts a drop-down list content control into the active document, sets the title and placeholder text and adds several items to the list, and then selects the last item entered.

## Properties (6)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified ContentControlListEntry object.
- `Text As String  (read/write)`  
  Returns or sets a String that represents the display text of a list item for a drop-down list or combo box content control. Read/write.
- `Value As String  (read/write)`  
  Returns or sets a String that represents the programmatic value of an item in a drop-down list or combo box content control. Read/write.
- `Index As Long  (read/write)`  
  Returns or sets a Long that represents the ordinal position of a content control list item in the collection of list items. Read/write.

## Methods (4)

- `Delete()`  
  Deletes the specified item in a combo box or drop-down list content control.
- `MoveUp()`  
  Moves an item in a drop-down list or combo box content control up one item, so that it is before the item that originally preceded it.
- `MoveDown()`  
  Moves an item in a drop-down list or combo box content control down one item, so that it is after the item that originally followed it.
- `Select()`  
  Selects the list entry in a drop-down list or combo box content control and sets the text of the content control to the value of the item.
