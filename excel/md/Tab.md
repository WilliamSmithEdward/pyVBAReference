# Tab

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024469-0000-0000-C000-000000000046}  

Represents the tab of a chart or a worksheet.

**Remarks:** Use the Tab property of the Chart object or Worksheet object to return a Tab object. After a Tab object is returned, you can use the ColorIndex property to determine the settings of a tab for a chart or worksheet.

**Example:**

```vba
Sub CheckTab()

 ' Determine if color index of 1st tab is set to none.
 If Worksheets(1).Tab.ColorIndex = xlColorIndexNone Then
 MsgBox "The color index is set to none for the first " & _
 "worksheet tab."
 Else
 MsgBox "The color index for the tab of the first worksheet " & _
 "is not set none."
 End If

End Sub
```

## Properties (7)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Color As Variant  (read/write)`  
  Returns or sets the primary color of the object, as shown in the table in the remarks section. Use the RGB function to create a color value. Read/write Variant.
- `ColorIndex As XlColorIndex  (read/write)`  
  Returns or sets a Variant value that represents the color of the specified worksheet tab or chart tab.
- `ThemeColor As XlThemeColor  (read/write)`  
  Returns or sets the theme color in the applied color scheme that is associated with the specified object. Read/write XlThemeColor.
- `TintAndShade As Variant  (read/write)`  
  Returns or sets a Single that lightens or darkens a color.
