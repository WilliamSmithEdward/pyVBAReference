# Parameters

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {0002442B-0000-0000-C000-000000000046}  

A collection of Parameter objects for the specified query table.

**Remarks:** Each Parameter object represents a single query parameter. Every query table contains a Parameters collection, but the collection is empty unless the query table is using a parameter query. You cannot use the Add method on a URL connection query table. For URL connection query tables, Microsoft Excel creates the parameters based on the Connection and PostText properties.

**Example:**

```vba
MsgBox Workbooks(1).ActiveSheet.QueryTables(1).Parameters.Count
```

## Properties (5)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Count As Long  (read-only)`  
  Returns a Long value that represents the number of objects in the collection.
- `_Default As Parameter  (read-only)`

## Methods (4)

- `Add(Name As String, [iDataType As Variant]) As Parameter`  
  Creates a new query parameter.
    - `Name As String` (required): The name of the specified parameter. The parameter name should match the parameter clause in the SQL statement.
    - `iDataType As Variant` (optional): The data type of the parameter. Can be any XlParameterDataType constant. These values correspond to ODBC data types. They indicate the type of value that the ODBC driver is expecting to receive. Microsoft Excel and the ODBC driver manager will coerce the parameter value given in Excel into the correct data type for the driver.
- `Item(Index As Variant) As Parameter`  
  Returns a single object from a collection.
    - `Index As Variant` (required): The name or index number for the object.
- `Delete()`  
  Deletes the object.
- `_NewEnum() As IUnknown`
