# ListLevels

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {0002098E-0000-0000-C000-000000000046}  

A collection of ListLevel objects that represents all the list levels of a list template, either the only level for a bulleted or numbered list or one of the nine levels of an outline numbered list.

**Remarks:** Use the ListLevels property to return the ListLevels collection. The following example sets the variable _mytemp_ to the first list template in the active document and then modifies each level to use lowercase letters for its number style. Use ListLevels (Index), where Index is a number from 1 through 9, to return a single ListLevel object. The following example sets list level one of list template one in the active document to start at four. To apply a list level, first identify the range or list, and then use the ApplyListTemplate method. Each tab at the beginning of the paragraph is translated into a list level. For example, a paragraph that begins with three tabs will become a level-three list paragraph after the ApplyListTemplate method is used.

## Properties (5)

- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of list levels in the collection. Read-only.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified ListLevels object.

## Methods (1)

- `Item(Index As Long) As ListLevel`  
  Returns an individual ListLevel object in a collection.
    - `Index As Long` (required): The individual object to be returned. Can be a Long indicating the ordinal position of the individual object.
