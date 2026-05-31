# Dialog

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {0002087A-0000-0000-C000-000000000046}  

Represents a built-in Microsoft Excel dialog box.

**Remarks:** The Dialog object is a member of the Dialogs collection. The Dialogs collection contains all the built-in dialog boxes in Microsoft Excel. You cannot create a new built-in dialog box or add one to the collection. The only useful thing that you can do with a Dialog object is to use it with the Show method to display the corresponding dialog box. The Microsoft Excel Visual Basic object library includes built-in constants for many of the built-in dialog boxes. Each constant is formed from the prefix "xlDialog" followed by the name of the dialog box. For example, the Apply Names dialog box constant is xlDialogApplyNames, and the Find File dialog box constant is xlDialogFindFile. These constants are members of the XlBuiltinDialog enumerated type.

**Example:**

```vba
dlgAnswer = Application.Dialogs(xlDialogOpen).Show
```

## Properties (3)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.

## Methods (1)

- `Show([Arg1 As Variant], [Arg2 As Variant], [Arg3 As Variant], [Arg4 As Variant], [Arg5 As Variant], [Arg6 As Variant], [Arg7 As Variant], [Arg8 As Variant], [Arg9 As Variant], [Arg10 As Variant], [Arg11 As Variant], [Arg12 As Variant], [Arg13 As Variant], [Arg14 As Variant], [Arg15 As Variant], [Arg16 As Variant], [Arg17 As Variant], [Arg18 As Variant], [Arg19 As Variant], [Arg20 As Variant], [Arg21 As Variant], [Arg22 As Variant], [Arg23 As Variant], [Arg24 As Variant], [Arg25 As Variant], [Arg26 As Variant], [Arg27 As Variant], [Arg28 As Variant], [Arg29 As Variant], [Arg30 As Variant]) As Boolean`  
  Displays the built-in dialog box, waits for the user to input data, and returns a Boolean value that represents the user's response.
