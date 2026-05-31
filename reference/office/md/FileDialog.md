# FileDialog

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C0362-0000-0000-C000-000000000046}  

Provides file dialog box functionality similar to the functionality of the standard Open and Save dialog boxes found in Microsoft Office applications.

**Remarks:** Use the FileDialog property to return a FileDialog object. The FileDialog property is located in each individual Office application's Application object. The property takes a single argument, MsoFileDialogType, that determines the type of FileDialog object that the property returns. There are four types of FileDialog object: - Open dialog box: lets users select one or more files that you can then open in the host application by using the Execute method. - SaveAs dialog box: lets users select a single file that you can then save the current file as by using the Execute method. - File Picker dialog box: lets users select one or more files. The file paths that the user selects are captured in the FileDialogSelectedItems collection. - Folder Picker dialog box: lets users select a path. The path that the user selects is captured in the FileDialogSelectedItems collection. Each host application can only create a single instance of the FileDialog object. Therefore, many of the properties of the FileDialog object persist even when you create multiple FileDialog objects. Therefore, make sure that you set all of the properties appropriately for your purpose before you display the dialog box.

**Example:**

```vba
Sub Main()

 'Declare a variable as a FileDialog object.
 Dim fd As FileDialog

 'Create a FileDialog object as a File Picker dialog box.
 Set fd = Application.FileDialog(msoFileDialogFilePicker)

 'Declare a variable to contain the path
 'of each selected item. Even though the path is aString,
 'the variable must be a Variant because For Each...Next
 'routines only work with Variants and Objects.
 Dim vrtSelectedItem As Variant

 'Use a With...End With block to reference the FileDialog object.
 With fd

 'Use the Show method to display the File Picker dialog box and return the user's action.
 'The user pressed the button.
 If .Show = -1 Then

 'Step through each string in the FileDialogSelectedItems collection.
 For Each vrtSelectedItem In .SelectedItems

 'vrtSelectedItem is aString that contains the path of each selected item.
 'Use any file I/O functions that you want to work with this path.
 'This example displays the path in a message box.
 MsgBox "The path is: " & vrtSelectedItem

 Next vrtSelectedItem
 'The user pressed Cancel.
 Else
 End If
 End With

 'Set the object variable to Nothing.
 Set fd = Nothing

End Sub
```

## Properties (13)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the FileDialog object (you can use this property with an Automation object to return that object's container application). Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the FileDialog object was created. Read-only.
- `Parent As Object  (read-only)`  
  Gets the Parent object for the FileDialog object. Read-only.
- `Filters As FileDialogFilters  (read-only)`  
  Gets a FileDialogFilters collection. Read-only.
- `FilterIndex As Long  (read/write)`  
  Gets or sets a Long indicating the default file filter of a file dialog box. The default filter determines which types of files are displayed when the file dialog box is first opened. Read/write.
- `Title As String  (read/write)`  
  Sets or gets the title of a file dialog box displayed by using the FileDialog object. Read/write.
- `ButtonName As String  (read/write)`  
  Sets or gets a String representing the text that is displayed on the action button of a file dialog box. Read/write.
- `AllowMultiSelect As Boolean  (read/write)`  
  Is True if the user is allowed to select multiple files from a file dialog box. Read/write.
- `InitialView As MsoFileDialogView  (read/write)`  
  Gets or sets an MsoFileDialogView constant representing the initial presentation of files and folders in a file dialog box. Read/write.
- `InitialFileName As String  (read/write)`  
  Sets or returns a String representing the path or file name that is initially displayed in a file dialog box. Read/write.
- `SelectedItems As FileDialogSelectedItems  (read-only)`  
  Gets a FileDialogSelectedItems collection. This collection contains a list of the paths of the files that a user selected from a file dialog box displayed by using the Show method of the FileDialog object. Read-only.
- `DialogType As MsoFileDialogType  (read-only)`  
  Gets an MsoFileDialogType constant representing the type of file dialog box that the FileDialog object is set to display. Read-only.
- `Item As String  (read-only)`  
  Gets the text associated with an object. Read-only.

## Methods (2)

- `Show() As Long`  
  Displays a file dialog box and returns a Long indicating whether the user pressed the Action button (-1) or the Cancel button (0). When you call the Show method, no more code executes until the user dismisses the file dialog box. In the case of the Open and SaveAs dialog boxes, use the Execute method right after the Show method to carry out the user's action.
- `Execute()`  
  Carries out a user's action right after the Show method is invoked.
