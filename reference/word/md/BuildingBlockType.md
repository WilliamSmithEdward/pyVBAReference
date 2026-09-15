# BuildingBlockType

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {817F99FA-CCC4-4971-8E9D-1238F735AAFF}  

Represents a type of building block. Each BuildingBlockType object is a member of the BuildingBlockTypes collection.

**Remarks:** Microsoft Word uses types and categories to organize building blocks. Each building block type is represented by a WdBuildingBlockTypes constant. Use the Categories property to access categories for a specific building block type. The following example prints the type and category names of all the building blocks in the first template to the Immediate Window. (This example assumes that the Immediate Window is visible.) For more information about building blocks, see Working with Building Blocks.

## Properties (6)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified BuildingBlockType object.
- `Index As Long  (read-only)`  
  Returns a Long that represents the position of an item in a collection. Read-only.
- `Name As String  (read-only)`  
  Returns a String that represents the localized name of a building block type. Read-only.
- `Categories As Categories  (read-only)`  
  Returns a Categories collection that represents the categories for a building block type. Read-only.
