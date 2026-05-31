# ModelMeasure

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244ED-0000-0000-C000-000000000046}  

Represents a single ModelMeasure object in the ModelMeasures collection.

## Properties (8)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Name As String  (read/write)`  
  The name of the model measure. Read/write String.
- `AssociatedTable As ModelTable  (read/write)`  
  Specifies the table that contains the model measure, as displayed in the Field List task pane. Read/write ModelTable.
- `Formula As String  (read/write)`  
  The Data Analysis Expressions (DAX) formula of the model measure. Read/write String.
- `FormatInformation As Variant  (read/write)`  
  The format of the model measure. Read/write Variant.
- `Description As String  (read/write)`  
  The description of the model measure. Read/write String.

## Methods (1)

- `Delete()`  
  Deletes the model measure from the data model.
