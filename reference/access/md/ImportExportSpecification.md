# ImportExportSpecification

**Type:** Dispatch Interface  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {55B0E0C9-C75D-4F42-AD20-6939C1D05B70}  

Represents a saved import or export operation.

**Remarks:** An ImportExportSpecification object contains all the information that Microsoft Access needs to repeat an import or export operation without your having to provide any input. For example, an import specification that imports data from a Microsoft Office Excel 2007 workbook stores the name of the source Excel file, the name of the destination database, and other details, such as whether you appended to or created a new table, primary key information, field names, and so on. Use the Add method of the ImportExportSpecifications collection to create a new ImportExportSpecification object. Use the Execute method to run a saved import or export operation.

## Properties (5)

- `Application As Application  (read-only)`  
  Use the Application property to access the active Microsoft Access Application object and its related properties. Read-only Application object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only Object.
- `Name As String  (read/write)`  
  Gets or sets the name of the specified ImportExportSpecification object. Read/write String.
- `XML As String  (read/write)`  
  Gets or sets the Extensible Markup Language (XML) string that defines an ImportExportSpecification object. Read/write String.
- `Description As String  (read/write)`  
  Gets or sets a String that describes the specified ImportExportSpecification object. Read/write.

## Methods (2)

- `Execute([Prompt As Variant])`  
  Executes the specified import or export specification.
- `Delete()`  
  Deletes the specified ImportExportSpecification object from the ImportExportSpecifications collection.
