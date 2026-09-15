# References

**Type:** Class  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {EB106214-9C89-11CF-A2B3-00A0C90542FF}  

The References collection contains Reference objects representing each reference that's currently set.

**Remarks:** The Reference objects in the References collection correspond to the list of references in the References dialog box, available by choosing References on the Tools menu. Each Reference object represents one selected reference in the list. References that appear in the References dialog box but haven't been selected aren't in the References collection. You can enumerate through the References collection by using the For Each...Next statement. The References collection belongs to the Microsoft Access Application object. Individual Reference objects in the References collection are indexed beginning with 1.

## Properties (2)

- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Count As Long  (read-only)`  
  Use the Count property to determine the number of items in a specified collection. Read-only Long.

## Methods (5)

- `Item(var As Variant) As Reference`  
  The Item method returns a specific member of a collection either by position or by key. Reference object.
    - `var As Variant` (required): An expression that specifies the position of a member of the collection. If a numeric expression, the var argument must be a number from 1 to the value of the collection's Count property. If a string expression, the var argument must be the name of a member of the collection.
- `_NewEnum() As IUnknown`
- `AddFromGuid(Guid As String, Major As Long, Minor As Long) As Reference`  
  The AddFromGUID method creates a Reference object based on the GUID that identifies a type library. Reference object.
    - `Guid As String` (required): A GUID that identifies a type library.
    - `Major As Long` (required): The major version number of the reference.
    - `Minor As Long` (required): The minor version number of the reference.
- `AddFromFile(FileName As String) As Reference`  
  The AddFromFile method creates a reference to a type library in a specified file.
    - `FileName As String` (required): A string expression that evaluates to the full path and file name of the file containing the type library to which you wish to set a reference.
- `Remove(Reference As Reference)`  
  The Remove method removes a Reference object from the References collection.
    - `Reference As Reference` (required): The Reference object that represents the reference that you wish to remove.

## Events (2)

- `ItemAdded(Reference As Reference)`  
  The ItemAdded event occurs when a reference is added to the project from Visual Basic.
    - `Reference As Reference` (required): The reference that was added to the project.
- `ItemRemoved(Reference As Reference)`  
  The ItemRemoved event occurs when a reference is removed from the project.
    - `Reference As Reference` (required): The reference that was removed from the project.
