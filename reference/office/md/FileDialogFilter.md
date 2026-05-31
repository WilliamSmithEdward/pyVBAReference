# FileDialogFilter

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C0364-0000-0000-C000-000000000046}  

Represents a file filter in a file dialog box displayed through the FileDialog object. Each file filter determines which files are displayed in the file dialog box.

**Remarks:** Use the Item method with the FileDialogFilters collection to return a FileDialogFilter object. Use the Add method to add a FileDialogFilter object to the FileDialogFilters collection. You can return the extensions that a FileDialogFilter object uses to filter files with the Extensions property, and you can return the description of the filter with the Description property; however, both of these properties are read-only. If you want to set the extension or description, you must use the Add method.

**Example:**

```vba
Sub Main()

 'Declare a variable as a FileDialogFilters collection.
 Dim fdfs As FileDialogFilters

 'Declare a variable as a FileDialogFilter object.
 Dim fdf As FileDialogFilter

 'Set the FileDialogFilters collection variable to
 'the FileDialogFilters collection of the SaveAs dialog box.
 Set fdfs = Application.FileDialog(msoFileDialogSaveAs).Filters

 'Iterate through the description and extensions of each
 'default filter in the SaveAs dialog box.
 For Each fdf In fdfs

 'Display the description of filters that include
 'Microsoft Excel files.
 If InStr(1, fdf.Extensions, "xls", vbTextCompare) > 0 Then
 MsgBox "Description of filter: " & fdf.Description
 End If
 Next fdf
End Sub
```

## Properties (5)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the FileDialogFilter object (you can use this property with an Automation object to return that object's container application). Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the FileDialogFilter object was created. Read-only.
- `Parent As Object  (read-only)`  
  Gets the Parent object for the FileDialogFilter object. Read-only.
- `Extensions As String  (read-only)`  
  Gets a value containing the extensions that determine which files are displayed in a file dialog box for each Filter object. Read-only.
- `Description As String  (read-only)`  
  Gets the description of each Filter object as a String value. The description is the text that is displayed in a file dialog box. Read-only.
