# ModelMeasures

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244EE-0000-0000-C000-000000000046}  

Represents a collection of ModelMeasure objects.

**Remarks:** Each measure is represented by a ModelMeasure object. Use the ModelMeasures property of the Model object to return the ModelMeasures collection.

## Properties (6)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Count As Long  (read-only)`  
  Returns an integer that represents the number of objects in the collection.
- `_Default As ModelMeasure  (read-only)`
- `_NewEnum As IUnknown  (read-only)`

## Methods (2)

- `Item(Index As Variant) As ModelMeasure`  
  Returns a single object from a collection.
    - `Index As Variant` (required): The name or index number of the object.
- `Add(MeasureName As String, AssociatedTable As ModelTable, Formula As String, FormatInformation As Variant, [Description As Variant]) As ModelMeasure`  
  Adds a model measure to the model.
    - `MeasureName As String` (required): The name of the model measure.
    - `AssociatedTable As ModelTable` (required): The model table associated with the model measure. This is the table that contains the model measure, as seen in the Field List task pane.
    - `Formula As String` (required): The Data Analysis Expressions (DAX) formula, inserted as a string.
    - `FormatInformation As Variant` (required): The formatting of the model measure. See Remarks.
    - `Description As Variant` (optional): The description associated with the model measure.
