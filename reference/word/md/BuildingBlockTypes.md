# BuildingBlockTypes

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {B5828B50-0E3D-448A-962D-A40702A5868D}  

Represents a collection of BuildingBlockType objects.

**Remarks:** Building block types are represented by WdBuildingBlockTypes constants. Use the Item method to access a specific type in the BuildingBlockTypes collection. To loop through the different building block types, use a For loop with the Count property. The following example loops through the building block types and prints the name in the Immediate Window. (This example assumes that the Immediate Window is visible.) For more information about building blocks, see Working with Building Blocks.

## Properties (4)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified BuildingBlockTypes object.
- `Count As Long  (read-only)`  
  Returns the number of items in the BuildingBlockTypes collection. Read-only Long.

## Methods (1)

- `Item(Index As WdBuildingBlockTypes) As BuildingBlockType`  
  Returns a BuildingBlockType object that represents the specified item in the collection.
    - `Index As WdBuildingBlockTypes` (required): Specifies the building block type of the item in the collection.
