# ModelRelationships

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244DA-0000-0000-C000-000000000046}  

This collection contains all relationships between data tables in the data model of Excel 2013.

## Properties (6)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Excel application. Read-only.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified ModelRelationships object. Read-only.
- `Count As Long  (read-only)`  
  Returns a Long value that represents the number of ModelRelationship objects in a ModelRelationships object. Read-only.
- `_Default As ModelRelationship  (read-only)`
- `_NewEnum As IUnknown  (read-only)`

## Methods (3)

- `Item(Index As Variant) As ModelRelationship`  
  Returns a single object from the ModelRelationships object.
    - `Index As Variant` (required): The index number for the object.
- `Add(ForeignKeyColumn As ModelTableColumn, PrimaryKeyColumn As ModelTableColumn) As ModelRelationship`  
  Adds a new relationship to the model.
    - `ForeignKeyColumn As ModelTableColumn` (required): A ModelTableColumn object representing the foreign key column in the table on the many side of the one-to-many relationship.
    - `PrimaryKeyColumn As ModelTableColumn` (required): A ModelTableColumn object representing the primary key column in the table on the one side of the one-to-many relationship.
- `DetectRelationships(PivotTable As PivotTable)`  
  Detects model relationships in the specified PivotTable object.
    - `PivotTable As PivotTable` (required): The PivotTable in which to detect model relationships.
