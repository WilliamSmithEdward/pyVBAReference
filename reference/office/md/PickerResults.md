# PickerResults

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C03E5-0000-0000-C000-000000000046}  

A collection of PickerResult objects.

**Remarks:** Each PickerResult object represents a resolved or selected item data.

**Example:**

```vba
Dim objPickerDialog As PickerDialog
Dim objPickerProperties As PickerProperties
Dim objPickerProperty As PickerProperty
Dim objPickerExistingResults As PickerResults
Dim objPickerExistingResults As PickerResult
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

' Enumerate the results.
For index = 1 To objPickerResults.Count-1
 Debug.Print objPickerResults.Item(index).Id
 Debug.Print objPickerResults.Item(index).DisplayName
 Debug.Print objPickerResults.Item(index).Type
 Debug.Print objPickerResults.Item(index).SIPId
Next
```

## Properties (5)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the PickerResults object. Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the PickerResults object was created. Read-only.
- `Item As PickerResult  (read-only)`  
  Retrieves a PickerResult object at the specified index. Read-only.
- `Count As Long  (read-only)`  
  Retrieves the count of the number of PickerResult objects contained within the PickerResults collection. Read-only.
- `_NewEnum As IUnknown  (read-only)`

## Methods (1)

- `Add(Id As String, DisplayName As String, Type As String, [SIPId As String], [ItemData As Variant], [SubItems As Variant]) As PickerResult`  
  Adds a PickerResult object to the PickerResults collection.
    - `Id As String` (required): Represents an identifier of the PickerResult.
    - `DisplayName As String` (required): Represents a display name of the PickerResult.
    - `Type As String` (required): Represents a type of the PickerResult.
    - `SIPId As String` (optional): Currently not supported. The SIPId is the identifier for Office Communication Server. It's used only for the people picking scenario.
    - `ItemData As Variant` (optional): Non-displaying item binding data.
    - `SubItems As Variant` (optional): Displays the purpose or non-display purpose field data of the PickerResult. It's used for passing column values in the PickerDialog.
