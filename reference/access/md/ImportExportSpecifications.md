# ImportExportSpecifications

**Type:** Dispatch Interface  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {77DC8648-F725-4371-88C3-6EB6C4894CA4}  

Represents the collection of available ImportExportSpecification objects.

**Remarks:** Use the Add method to create a new ImportExportSpecification object. Use the ImportExportSpecification property of the CodeProject or CurrentProject object to return the ImportExportSpecifications collection.

## Properties (4)

- `Application As Application  (read-only)`  
  Use the Application property to access the active Microsoft Access Application object and its related properties. Read-only Application object.
- `Count As Long  (read-only)`  
  Use the Count property to determine the number of items in a specified collection. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only Object.
- `Item As ImportExportSpecification  (read-only)`  
  The Item property returns a specific member of a collection either by position or by index. Read-only ImportExportSpecification.

## Methods (1)

- `Add(Name As String, SpecificationDefinition As String) As ImportExportSpecification`  
  Adds a new ImportExportSpecification object to the ImportExportSpecifications collection.
    - `Name As String` (required): The name to use for the ImportExportSpecification.
    - `SpecificationDefinition As String` (required): The XML markup that represents the settings to save for the import or export operation.
