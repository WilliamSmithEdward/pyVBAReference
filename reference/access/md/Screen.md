# Screen

**Type:** Dispatch Interface  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {DC6B66C0-6128-101B-AF4E-00AA003F0F07}  

The Screen object refers to the particular form, report, or control that currently has the focus.

**Remarks:** Use the Screen object together with its properties to refer to a particular form, report, or control that has the focus. For example, you can use the Screen object with the ActiveForm property to refer to the form in the active window without knowing the form's name. The following example displays the name of the form in the active window. Referring to the Screen object doesn't make a form, report, or control active. To make a form, report, or control active, you must use the SelectObject method of the DoCmd object. If you refer to the Screen object when there's no active form, report, or control, Microsoft Access returns a run-time error. For example, if a standard module is in the active window, the code in the preceding example would return an error.

**Example:**

```vba
Sub ActiveObjects()
 Dim frm As Form, ctl As Control

 ' Return Form object pointing to active form.
 Set frm = Screen.ActiveForm
 MsgBox frm.Name & " is the active form."
 ' Return Control object pointing to active control.
 Set ctl = Screen.ActiveControl
 MsgBox ctl.Name & " is the active control " _
 & "on this form."
End Sub
```

## Properties (8)

- `Application As Application  (read-only)`  
  Use the Application property to access the active Microsoft Access Application object and its related properties. Read-only Application object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `ActiveDatasheet As Form  (read-only)`  
  Use the ActiveDatasheet property together with the Screen object to identify or refer to the datasheet that has the focus. Read-only Form object.
- `ActiveControl As Control  (read-only)`  
  Use the ActiveControl property together with the Screen object to identify or refer to the control that has the focus. Read-only Control object.
- `PreviousControl As Control  (read-only)`  
  Use the PreviousControl property together with the Screen object to return a reference to the control that last received the focus. Read-only.
- `ActiveForm As Form  (read-only)`  
  Use the ActiveForm property together with the Screen object to identify or refer to the form that has the focus. Read-only Form object.
- `ActiveReport As Report  (read-only)`  
  Use the ActiveReport property together with the Screen object to identify or refer to the report that has the focus. Read-only Report object.
- `MousePointer As Integer  (read/write)`  
  Use the MousePointer property together with the Screen object to specify or determine the type of mouse pointer currently displayed. Read/write Integer.
