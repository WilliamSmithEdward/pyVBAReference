# ContentControlListEntries

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {54F46DC4-F6A6-48CC-BD66-46C1DDEADD22}  

The ContentControlListEntries collection contains ContentControlListEntry objects that represent the items in a drop-down list or combo box content control.

**Remarks:** Use the Add method to add an item to a drop-down list or combo box. The following code example uses the Add method to add several list items to a new drop-down list content control. Use the Item method or the DropdownListEntries property of a ContentControl object to access an individual list item within a collection. The following code example uses the Item method to access the third item in a list and change the display text. Use the Clear method to remove all items from a drop-down list or combo box. The following code example clears all items from the first content control in the active document.

## Properties (5)

- `_NewEnum As IUnknown  (read-only)`
- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified ContentControlListEntries object.
- `Count As Long  (read-only)`  
  Returns the number of items in the ContentControlListEntries collection. Read-only Long.

## Methods (3)

- `Clear()`  
  Clears all items from a drop-down list or combo box content control.
- `Item(Index As Long) As ContentControlListEntry`  
  Returns a ContentControlListEntry object that represents the specified item in the collection.
    - `Index As Long` (required): Specifies the ordinal position of the object within the collection.
- `Add(Text As String, [Value As String], [Index As Long]) As ContentControlListEntry`  
  Adds a new list item to a drop-down list or combo box content control and returns a ContentControlListEntry object.
    - `Text As String` (required): Specifies the display text for the list item. Corresponds to the Text property for a ContentControlListEntry object.
    - `Value As String` (optional): Specifies the value of the list item. Corresponds to the Value property for a ContentControlListEntry object. If omitted, the Value property is equal to the Text property.
    - `Index As Long` (optional): Specifies the ordinal position of the new item in the list. If an item exists at the position specified, the existing item is pushed down in the list. If omitted, the new item is added to the end of the list.
