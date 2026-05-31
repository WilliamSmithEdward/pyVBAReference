# PickerProperties

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C03E3-0000-0000-C000-000000000046}  

A collection of PickerProperty objects.

**Remarks:** Each PickerProperty object is a Name(ID)/Value pair for passing option values to a PickerDialog object. You can get a PickerProperties collection object through the Properties property of the PickerDialog object.

**Example:**

```vba
Dim objPickerDialog As PickerDialog
Dim objPickerProperties As PickerProperties
Dim objPickerProperty As PickerProperty
Dim objPickerExistingResults As PickerResults
Dim objPickerExistingResult As PickerResult
Dim objPickerResults As PickerResults


' Configure the Picker Dialog properties.
Set objPickerDialog = Application.PickerDialog
objPickerDialog.DataHandlerId = "{000CDF0A-0000-0000-C000-000000000046}"
objPickerDialog.Title = "Sample Picker Dialog"
Set objPickerProperties = objPickerDialog.Properties
Set objPickerProperty = objPickerProperties.Add("SiteUrl", "https://my", msoPickerFieldtypeText)
Set objPickerExistingResults = objPickerDialog.CreatePickerResults
Set objPickerExistingResult = objPickerExistingResults.Add("johndoe@contoso.com", "John Doe", "User")

' Show the Picker Dialog and get the results.
Set objPickerResults = objPickerDialog.Show(True, objPickerExistingResult)
```

## Properties (5)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the PickerProperties object. Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the PickerProperties object was created. Read-only.
- `Item As PickerProperty  (read-only)`  
  Retrieves a PickerProperty object at the specified index. Read-only.
- `Count As Long  (read-only)`  
  Retrieves the count of the number of PickerProperty objects contained within the PickerProperties collection. Read-only.
- `_NewEnum As IUnknown  (read-only)`

## Methods (2)

- `Add(Id As String, Value As String, Type As MsoPickerField) As PickerProperty`  
  Adds a PickerProperty object to the collection.
    - `Id As String` (required): Key name of the property.
    - `Value As String` (required): Value of the property.
    - `Type As MsoPickerField` (required): Type of the property.
- `Remove(Id As String)`  
  Removes a PickerProperty object from the collection.
    - `Id As String` (required): The identifier of the PickerProperty object to remove.
