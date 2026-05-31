# ProtectedViewWindow

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244CD-0000-0000-C000-000000000046}  

Represents a Protected View window.

**Remarks:** A Protected View window is used to display a workbook from a potentially unsafe location. Unsafe locations are defined as the following: - Files opened from the Internet. - Attachments opened from Outlook. - Files blocked by File Block Policy. - Files that fail Office file validation. - Files explicitly opened in Protected View by using the Open in Protected View command of the Open button in the Open dialog box. Workbooks displayed in a Protected View window cannot be edited and are restricted from running active content such as Visual Basic for Applications macros and data connections. For more information about Protected View windows, see What is Protected View? To return a single ProtectedViewWindow object from the ProtectedViewWindows collection, use ProtectedViewWindows (_index_), where _index_ is the index number of the window that you want to open. You can also access the ProtectedViewWindow object that represents the active Protected View window by using the ActiveProtectedViewWindow property of the Application object.

**Example:**

```vba
Dim wbProtected As Workbook

If Application.ProtectedViewWindows.Count > 0 Then
    Set wbProtected = Application.ProtectedViewWindows(1).Workbook
End If
```

## Properties (12)

- `_Default As String  (read-only)`
- `Caption As String  (read/write)`  
  Returns or sets a Variant value that represents the name that appears in the title bar of the Protected View window. Read/write.
- `EnableResize As Boolean  (read/write)`  
  True if the Protected View window can be resized. Read/write.
- `Height As Double  (read/write)`  
  Returns or sets a value that represents the height, in points, of the Protected View window. Read/write.
- `Left As Double  (read/write)`  
  Returns or sets a value that represents the distance, in points, from the left edge of the client area to the left edge of the Protected View window. Read/write.
- `Top As Double  (read/write)`  
  Returns or sets a value that represents the distance, in points, from the top edge of the specified Protected View window to the top edge of the usable area. Read/write.
- `Width As Double  (read/write)`  
  Returns or sets a value that specifies the width, in points, of the specified Protected View window. Read/write.
- `Visible As Boolean  (read/write)`  
  Returns or sets a value that determines whether the specified Protected View window is visible. Read/write.
- `SourceName As String  (read-only)`  
  Returns the name of the source file that is open in the specified Protected View window. Read-only.
- `SourcePath As String  (read-only)`  
  Returns the path of the source file that is open in the specified Protected View window. Read-only.
- `WindowState As XlProtectedViewWindowState  (read/write)`  
  Returns or sets the state of the specified Protected View window. Read/write.
- `Workbook As Workbook  (read-only)`  
  Returns an object that represents the workbook that is open in the specified Protected View window. Read-only.

## Methods (3)

- `Activate()`  
  Brings the Protected View window to the front of the z-order.
- `Close() As Boolean`  
  Closes the specified Protected View window.
- `Edit([WriteResPassword As Variant], [UpdateLinks As Variant]) As Workbook`  
  Opens the workbook that is open for editing in the specified Protected View window.
    - `WriteResPassword As Variant` (optional): The password required to write to a write-reserved workbook. If this argument is omitted and the workbook requires a password, the user will be prompted for the password.
    - `UpdateLinks As Variant` (optional): Specifies the way that external references (links) in the file, such as the reference to a range in the Budget.xls workbook in the following formula =SUM([Budget.xls]Annual!C10:C25), are updated. If this argument is omitted, the user is prompted to specify how links will be updated. For more information about the values used by this parameter, see the Remarks section. If Excel is opening a file in the WKS, WK1, or WK3 format and the _UpdateLinks_ argument is 0, no charts are created; otherwise, Excel generates charts from the graphs attached to the file.
