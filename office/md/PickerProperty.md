# PickerProperty

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C03E2-0000-0000-C000-000000000046}  

Represents an object for passing a custom property.

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
  Gets an Application object that represents the container application for the PickerProperty object. Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the PickerProperty object was created. Read-only.
- `Id As String  (read-only)`  
  Retrieves the unique Id of the associated PickerProperty object. Read-only.
- `Value As Variant  (read-only)`  
  Retrieves the value of a PickerProperty. Read-only.
- `Type As MsoPickerField  (read-only)`  
  Retrieves the type of the PickerProperty. Read-only.
