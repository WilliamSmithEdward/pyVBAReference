# NewFile

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C0936-0000-0000-C000-000000000046}  

Represents items listed on the New Item task pane available in several Microsoft Office applications.

**Remarks:** Use the Application property or the Creator property to access the NewFile object in each of the applications.

**Example:**

```vba
Sub AddNewDocToTaskPane()
    Application.NewDocument.Add FileName:="C:\NewDocument.doc", _
        Section:=msoNew, DisplayName:="New Document"
    CommandBars("Task Pane").Visible = True
End Sub
```

## Properties (2)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the NewFile object (you can use this property with an Automation object to return that object's container application). Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the NewFile object was created. Read-only.

## Methods (2)

- `Add(FileName As String, [Section As Variant], [DisplayName As Variant], [Action As Variant]) As Boolean`  
  Adds a new item to the New Item task pane. Returns a Boolean value to indicate whether the operation was successful.
    - `FileName As String` (required): The name of the file to add to the list of files on the task pane.
    - `Section As Variant` (optional): The section to which to add the file. Can be any msoFileNew constant.
    - `DisplayName As Variant` (optional): The text to display in the task pane.
    - `Action As Variant` (optional): The action to take when a user clicks the item. Can be any msoFileNew constant.
- `Remove(FileName As String, [Section As Variant], [DisplayName As Variant], [Action As Variant]) As Boolean`  
  Removes an item from the New Item task pane. Returns a Boolean value to indicate whether the operation was successful.
    - `FileName As String` (required): The name of the file reference.
    - `Section As Variant` (optional): The section of the task pane where the file reference exists. Can be any msoFileNew constant.
    - `DisplayName As Variant` (optional): The display text of the file reference.
    - `Action As Variant` (optional): The action to take when a user clicks the item. Can be any msoFileNew constant.
