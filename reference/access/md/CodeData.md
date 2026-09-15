# CodeData

**Type:** Class  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {08F6C81A-3CFD-11D1-98BC-006008197D41}  

The CodeData object refers to objects stored within the code database by the source (server) application.

**Remarks:** The CodeData object has several collections that contain specific object types within the code database. The following table lists the name of each collection defined by the database and the types of objects it contains. For example, an AccessObject representing a table is a member of the AllTables collection, which is a collection of AccessObject objects within the current database. Within the AllTables collection, individual tables are indexed beginning with zero. Refer to an individual AccessObject object in the AllTables collection either by referring to the table by name, or by referring to its index within the collection. If you want to refer to a specific item in the AllTables collection, it's better to refer to it by name because the item's index may change. If the object name includes a space, the name must be surrounded by brackets ([ ]).

## Properties (6)

- `AllTables As AllTables  (read-only)`  
  Use the AllTables property to reference the AllTables collection and its related properties. Read-only AllTables object.
- `AllQueries As AllQueries  (read-only)`  
  Use the AllQueries property to reference the AllQueries collection and its related properties. Read-only AllQueries object.
- `AllViews As AllViews  (read-only)`  
  Use the AllViews property to reference the AllViews collection and its related properties. Read-only AllViews object.
- `AllStoredProcedures As AllStoredProcedures  (read-only)`  
  Use the AllStoredProcedures property to reference the AllStoredProcedures collection and its related properties. Read-only AllStoredProcedures object.
- `AllDatabaseDiagrams As AllDatabaseDiagrams  (read-only)`  
  Use the AllDatabaseDiagrams property to reference the AllDatabaseDiagrams collection and its related properties. Read-only AllDatabaseDiagrams object.
- `AllFunctions As AllFunctions  (read-only)`  
  Use the AllFunctions property to reference the AllFunctions collection and its related properties. Read-only AllFunctions object.
