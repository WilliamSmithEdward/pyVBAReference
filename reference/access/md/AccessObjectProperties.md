# AccessObjectProperties

**Type:** Dispatch Interface  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {0921F331-A7C9-11D1-9944-006008197D41}  

The AccessObjectProperties collection contains all of the custom AccessObjectProperty objects of a specific instance of an object. These AccessObjectProperty objects (which are often just called properties) uniquely characterize that instance of the object.

**Remarks:** Use the AccessObjectProperties collection in Visual Basic or in an expression to refer to properties of the CurrentProject, CodeProject, or AccessObject object. For example, you can enumerate the AccessObjectProperties collection to set or return the values of properties of an individual report. To add a user-defined property to an existing instance of an object, first define its characteristics and add it to the collection with the Add method. Referencing a user-defined AccessObjectProperty object that has not yet been appended to an AccessObjectProperties collection will cause an error, as will appending a user-defined AccessObjectProperty object to an AccessObjectProperties collection containing an AccessObjectProperty object of the same name. Use the Remove method to remove user-defined properties from the AccessObjectProperties collection. To refer to a built-in or user-defined AccessObjectProperty object in a collection by its ordinal number or by its Name property setting, use any of the following syntax forms. With the same syntax forms, you can also refer to the Value property of an AccessObjectProperty object.

## Properties (4)

- `Application As Application  (read-only)`  
  Use the Application property to access the active Microsoft Access Application object and its related properties. Read-only Application object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Item As AccessObjectProperty  (read-only)`  
  The Item property returns a specific member of a collection either by position or by index. Read-only AccessObjectProperty.
- `Count As Long  (read-only)`  
  Use the Count property to determine the number of items in a specified collection. Read-only Long.

## Methods (2)

- `Add(PropertyName As String, Value As Variant)`  
  Use the Add method to add a new property as an AccessObjectProperty object to the AccessObjectProperties collection of an AccessObject object.
    - `PropertyName As String` (required): A string expression that's the name of the new property.
    - `Value As Variant` (required): A Variant value corresponding to the option setting. The setting of the value argument depends on the possible settings for a particular option. Can be a constant or a string value.
- `Remove(Item As Variant)`  
  Use the Remove method to remove an AccessObjectProperty object from the AccessObjectProperties collection of an AccessObject object.
    - `Item As Variant` (required): An expression that specifies the position of a member of the collection referred to by the object argument. If a numeric expression, the index argument must be a number from 0 to the value of the collection's Count property minus 1. If a string expression, the index argument must be the name of a member of the collection.
