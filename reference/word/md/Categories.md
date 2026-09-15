# Categories

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {6E47678B-A879-4E56-8698-3B7CF169FAD4}  

Represents a collection of building block categories.

**Remarks:** Use the Item method to access an existing category. You can then use the BuildingBlocks property to access a collection of BuildingBlock objects for the category. The following example prints the type and category names of all the building blocks in the first template to the Immediate Window. (This example assumes that the Immediate Window is visible.) Use the Item method to access an existing category; to create a new category, use the Add method of the BuildingBlockEntries collection. Set the value of the Category parameter. For more information about building blocks, see Working with Building Blocks.

## Properties (4)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Categories collection.
- `Count As Long  (read-only)`  
  Returns the number of items in the Categories collection. Read-only Long.

## Methods (1)

- `Item(Index As Variant) As Category`  
  Returns a Category object that represents the specified item in the collection.
    - `Index As Variant` (required): Specifies a String or Integer that represents the name or ordinal position of the object in the collection.
