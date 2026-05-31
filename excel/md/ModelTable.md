# ModelTable

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244D7-0000-0000-C000-000000000046}  

Represents a table in the data model.

**Remarks:** The ModelTable object is read-only (cannot be created or edited programmatically). There is a ModelTable object for every table in the model.

## Properties (8)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Excel application. Read-only.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified ModelTable object. Read-only.
- `Name As String  (read-only)`  
  Returns a String value that represents the name of the object.
- `SourceName As String  (read-only)`  
  Name of the data source for the table. If the table has no data source, the call will generate a run-time error. Read-only String.
- `ModelTableColumns As ModelTableColumns  (read-only)`  
  Collection of ModelTableColumns objects that make up the ModelTable. Read-only.
- `SourceWorkbookConnection As WorkbookConnection  (read-only)`  
  Returns the Workbook Connection from which the model table originated. Read-only.
- `RecordCount As Long  (read-only)`  
  Returns the total row count for the model table. Read-only Integer.

## Methods (1)

- `Refresh()`  
  Refreshes the model table source connections.
