# ModelChanges

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244E4-0000-0000-C000-000000000046}  

Represents changes made to the data model.

**Remarks:** The ModelChanges object contains information about which changes were made to the data model when the ModelChange event of the Workbook object occurs after a model operation. When Micrososft Excel makes changes to the data model, multiple changes can be made in the same operation, and the ModelChanges object will include information about all the changes made in one model operation.

## Properties (14)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Excel application. Read-only.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified ModelChanges object. Read-only.
- `TablesAdded As ModelTableNames  (read-only)`  
  Returns a ModelTableNames collection of table names as strings representing all tables that were added to the model as part of a model operation. Read-only.
- `TablesDeleted As ModelTableNames  (read-only)`  
  Returns a ModelTableNames collection of table names as strings representing all tables that were deleted from the model as part of a model operation. Read-only.
- `TablesModified As ModelTableNames  (read-only)`  
  Returns a ModelTableNames collection of table names as strings representing all tables that were refreshed or recalculated as part of a model operation. Read-only.
- `TableNamesChanged As ModelTableNameChanges  (read-only)`  
  Returns a ModelTableNameChanges collection of ModelTableNameChange objects representing old and new names of all tables that were renamed in the model as part of a model operation. Read-only.
- `RelationshipChange As Boolean  (read-only)`  
  When True, one or more relationships in the model were changed (added, deleted or modified) as part of a model operation. When False, no relationships were changed during the operation. Read-only Boolean.
- `ColumnsAdded As ModelColumnNames  (read-only)`  
  Returns a ModelColumnNames collection of ModelColumnName objects that represent all columns added as part of a model operation. Read-only.
- `ColumnsDeleted As ModelColumnNames  (read-only)`  
  Returns a ModelColumnNames collection of ModelColumnName objects that represent all columns that were deleted as part of a model operation. Read-only.
- `ColumnsChanged As ModelColumnChanges  (read-only)`  
  Returns a ModelColumnChanges collection of ModelColumnChange objects that represent table names and column names of all table columns for which the data type was changed as part of a model operation. Read-only.
- `MeasuresAdded As ModelMeasureNames  (read-only)`  
  Returns a ModelMeasureNames collection of ModelMeasureName objects that represent all measures that were added as part of a model operation. Read-only.
- `UnknownChange As Boolean  (read-only)`  
  True when a non-specified change was made to the model as part of a model transaction. Read-only Boolean.
- `Source As XlModelChangeSource  (read-only)`  
  Returns the source of the data model. Read-only.
