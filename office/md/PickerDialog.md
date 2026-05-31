# PickerDialog

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C03E6-0000-0000-C000-000000000046}  

Provides dialog user interface functionality for picking people or picking data.

**Remarks:** Get the PickerDialog object through the PickerDialog property in the Application object.

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
  Gets an Application object that represents the container application for the PickerDialog object. Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the PickerDialog object was created. Read-only.
- `DataHandlerId As String  (read/write)`  
  Sets or gets the GUID of the PickerDialog data handler component. Read/write.
- `Title As String  (read/write)`  
  Sets or returns the title of a PickerDialog displayed in the PickerDialog. Read/write.
- `Properties As PickerProperties  (read-only)`  
  Returns the PickerProperties object to specify custom properties for the data handler component. Read-only.

## Methods (3)

- `CreatePickerResults() As PickerResults`  
  Creates an empty PickerResults object.
- `Show([IsMultiSelect As Boolean], [ExistingResults As PickerResults]) As PickerResults`  
  Displays the PickerDialog with the already specified data handler and given options.
    - `IsMultiSelect As Boolean` (optional): Specifies whether the PickerDialog user interface provides multiple item selection functions.
    - `ExistingResults As PickerResults` (optional): Contains existing PickerResults in the PickerDialog user interface. These results are displayed in the selected item control.
- `Resolve(TokenText As String, duplicateDlgMode As Long) As PickerResults`  
  Resolves the token by using the PickerDialog and retrieves the results.
    - `TokenText As String` (required): The text string to resolve.
