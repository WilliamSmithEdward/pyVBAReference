# Category

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {ECFBDB5E-ACD2-4530-AD79-4560B7FF055C}  

Represents an individual category of a building block type.

**Remarks:** Microsoft Word uses types and categories to organize building blocks. Each building block type is represented by a WdBuildingBlockTypes constant. Each category is a unique string that a user defines. Word comes with two categories already defined: "General" and "Custom"; you can create additional categories as you need. Use the Type property to access the building block type associated with a specific category. Use the BuildingBlocks property to access the collection of building blocks for a category. The following example prints the type and category names of all the building blocks in the first template to the Immediate Window. (This example assumes that the Immediate Window is visible.) Use the Item method of the Categories collection to access an existing category; to create a new category, use the Add method of the BuildingBlockEntries collection. Set the value of the Category parameter. For more information about building blocks, see Working with Building Blocks.

## Properties (7)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Category object.
- `Index As Long  (read-only)`  
  Returns a Long that represents the position of an item in a collection. Read-only.
- `Name As String  (read-only)`  
  Returns the name of the specified object. Read-only String.
- `BuildingBlocks As BuildingBlocks  (read-only)`  
  Returns a BuildingBlocks collection that represents the building blocks for a category. Read-only.
- `Type As BuildingBlockType  (read-only)`  
  Returns a BuildingBlockType object that represents the type of building block for a building block category. Read-only.
