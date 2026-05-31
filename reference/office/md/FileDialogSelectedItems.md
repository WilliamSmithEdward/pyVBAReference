# FileDialogSelectedItems

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C0363-0000-0000-C000-000000000046}  

A collection of String values that correspond to the paths of the files or folders that a user has selected from a file dialog box displayed through the FileDialog object.

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

 'Allow the selection of multiple file.
 .AllowMultiSelect = True

 'Use the Show method to display the File Picker dialog box and return the user's action.
 'The user pressed the button.
 If .Show = -1 Then

 'Step through each string in the FileDialogSelectedItems collection
 For Each vrtSelectedItem In .SelectedItems

 'vrtSelectedItem is aString that contains the path of each selected item.
 'Use any file I/O functions that you want to work with this path.
 'This example displays the path in a message box.
 MsgBox "Selected item's path: " & vrtSelectedItem

 Next vrtSelectedItem
 'The user pressed Cancel.
 Else
 End If
 End With

 'Set the object variable to Nothing.
 Set fd = Nothing

End Sub
```

## Properties (5)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the FileDialogSelectedItems object (you can use this property with an Automation object to return that object's container application). Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the FileDialogSelectedItems object was created. Read-only.
- `Parent As Object  (read-only)`  
  Gets the Parent object for the FileDialogSelectedItems object. Read-only.
- `_NewEnum As IUnknown  (read-only)`
- `Count As Long  (read-only)`  
  Gets a Long indicating the number of items in the FileDialogSelectedItems collection. Read-only.

## Methods (1)

- `Item(Index As Long) As String`  
  Gets a String that corresponds to the path of one of the files by the _Index_ that the user selected from a file dialog box that was displayed by using the Show method of the FileDialog object.
    - `Index As Long` (required): The index is one-based.
