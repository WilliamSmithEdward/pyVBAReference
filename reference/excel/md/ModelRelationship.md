# ModelRelationship

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244D9-0000-0000-C000-000000000046}  

Represents a relationship, currently in the data model, between two tables.

## Properties (8)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Excel application. Read-only.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified ModelRelationship object. Read-only.
- `ForeignKeyTable As ModelTable  (read-only)`  
  Contains the ModelTable object representing the table on the many side of the one-to-many relationship. Read-only.
- `ForeignKeyColumn As ModelTableColumn  (read-only)`  
  Contains the ModelTableColumn object representing the foreign key column on the many side of the one-to-many relationship. Read-only.
- `PrimaryKeyTable As ModelTable  (read-only)`  
  Contains the ModelTable object representing the table on the one side of the one-to-many relationship.
- `PrimaryKeyColumn As ModelTableColumn  (read-only)`  
  Contains the ModelTableColumn object representing the primary key column in the table on the one side of the one-to-many relationship.
- `Active As Boolean  (read/write)`  
  When True, the relationship is active. When False, this relationship is inactive. Read/write Boolean.

## Methods (1)

- `Delete()`  
  Deletes the relationship.
