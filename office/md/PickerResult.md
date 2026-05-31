# PickerResult

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C03E4-0000-0000-C000-000000000046}  

Represents a resolved or selected item of data.

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

## Properties (10)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the PickerResult object. Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the PickerResult object was created. Read-only.
- `Id As String  (read-only)`  
  Retrieves the unique Id of the associated PickerResult object. Read-only.
- `DisplayName As String  (read/write)`  
  Represents a display name of PickerResult. Read/write.
- `Type As String  (read/write)`  
  Represents the type of a PickerResult object. Read/write.
- `SIPId As String  (read/write)`  
  The identifier for Office Communication Server, used only for the people picking scenario. Read/write.
- `ItemData As Variant  (read/write)`  
  Gets or sets a non-display purpose item binding to data. Read/write.
- `SubItems As Variant  (read/write)`  
  Display purpose or non-display purpose field data of a PickerResult object, used for passing column values in a PickerDialog. Read/write.
- `DuplicateResults As Variant  (read-only)`  
  Gets PickerResult collection if the result of resolving results has multiple candidates. Read-only.
- `Fields As PickerFields  (read/write)`  
  Represents field definitions of SubItems in a PickerFields collection. Read-only.
