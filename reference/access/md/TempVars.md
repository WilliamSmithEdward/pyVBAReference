# TempVars

**Type:** Class  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {6D40D9DE-2821-44A8-BAF3-8011E362CF59}  

Represents the collection of TempVar objects.

**Remarks:** Use the Add method or the SetTempVar macro action to create a TempVar object. Use the Remove method or the RemoveTempVar macro action to delete a TempVar object from the TempVars collection. Use the RemoveAll method or the RemoveAllTempVars macro action to delete all TempVar objects from the TempVars collection. The TempVars collection can store up to 255 TempVar objects. If you don't remove a TempVar object, it will remain in memory until you close the database. It's a good practice to remove TempVar object variables when you are finished using them. To refer to a TempVar object in a collection by its ordinal number or by its Name property setting, use the following syntax form: - TempVar![name]

## Properties (4)

- `Application As Application  (read-only)`  
  Use the Application property to access the active Microsoft Access Application object and its related properties. Read-only Application object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only Object.
- `Item As TempVar  (read-only)`  
  The Item property returns a specific member of a collection either by position or by index. Read-only TempVar.
- `Count As Long  (read-only)`  
  Gets the number of TempVar objects in the TempVars collection. Read-only Long.

## Methods (3)

- `Add(Name As String, Value As Variant)`  
  Adds a variable to the TempVars collection.
    - `Name As String` (required): The name to use for the TempVar.
    - `Value As Variant` (required): The value to store as a TempVar. This value must be a string expression or a numeric expression. Setting this argument to an object data type will result in a run-time error.
- `Remove(var As Variant)`  
  Removes the specified TempVar object from the TempVars collection.
    - `var As Variant` (required): An expression that specifies the position of a member of the collection referred to by the expression argument. If a numeric expression, the argument must be a number from 0 to the value of the collection's Count property minus 1. If a string expression, the argument must be the name of a member of the collection.
- `RemoveAll()`  
  Removes all of the TempVar objects from the TempVars collection.
