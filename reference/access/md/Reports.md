# Reports

**Type:** Dispatch Interface  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {D1523700-6128-101B-AF4E-00AA003F0F07}  

The Reports collection contains all of the currently open reports in a Microsoft Access database.

**Remarks:** Use the Reports collection in Visual Basic or in an expression to refer to reports that are currently open. For example, you can enumerate the Reports collection to set or return the values of properties of individual reports in the collection. Refer to an individual Report object in the Reports collection either by referring to the report by name, or by referring to its index within the collection. The Reports collection is indexed beginning with zero. If you refer to a report by its index, the first report is Reports(0), the second report is Reports(1), and so on. If you opened Report1 and then opened Report2, Report2 would be referenced in the Reports collection by its index as Reports(1). If you then closed Report1, Report2 would be referenced in the Reports collection by its index as Reports(0). You can't add or delete a Report object from the Reports collection.

## Properties (4)

- `Application As Application  (read-only)`  
  Use the Application property to access the active Microsoft Access Application object and its related properties. Read-only Application object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Item As Report  (read-only)`  
  The Item property returns a specific member of a collection either by position or by index. Read-only Report.
- `Count As Long  (read-only)`  
  Use the Count property to determine the number of items in a specified collection. Read-only Long.
