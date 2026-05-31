# FileDialogFilters

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C0365-0000-0000-C000-000000000046}  

A collection of FileDialogFilter objects that represent the types of files that can be selected in a file dialog box that is displayed by using the FileDialog object.

**Example:**

```vba
Application.FileDialog(msoFileDialogOpen).Filters
```

## Properties (5)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the FileDialogFilters object (you can use this property with an Automation object to return that object's container application). Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the FileDialogFilters object was created. Read-only.
- `Parent As Object  (read-only)`  
  Gets the Parent object for the FileDialogFilters object. Read-only.
- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Gets a Long indicating the number of items in the FileDialogFilters collection. Read-only.

## Methods (4)

- `Item(Index As Long) As FileDialogFilter`  
  Gets a FileDialogFilter object that is a member of the specified FileDialogFilters collection.
- `Delete([filter As Variant])`  
  Removes a file dialog filter.
    - `filter As Variant` (optional): The filter to be removed.
- `Clear()`  
  Removes all filters currently applied in a file dialog box.
- `Add(Description As String, Extensions As String, [Position As Variant]) As FileDialogFilter`  
  Adds a new file filter to the list of filters in the Files of type drop-down list in the File dialog box. Returns a FileDialogFilter object that represents the newly added file filter.
    - `Description As String` (required): The text representing the description of the file name extension that you want to add to the list of filters.
    - `Extensions As String` (required): The text representing the file name extension you want to add to the list of filters. More than one extension may be specified, and each must be separated by a semi-colon. For example, the argument can be assigned to the string: ".txt; .htm". NOTE: Parentheses don't need to be added around the extensions. Microsoft Office will automatically add parentheses around the extensions string when the description and extensions strings are concatenated into one file filter item.
    - `Position As Variant` (optional): A number that indicates the position of the new control in the filter list. The new filter will be inserted before the filter at this position. If this argument is omitted, the filter is added at the end of the list.
