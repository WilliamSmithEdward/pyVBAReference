# Control

**Type:** Class  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {02F92C80-8F8E-101B-AF4E-00AA003F0F07}  

The Control object represents a control on a form, report, or section, within another control, or attached to another control.

**Remarks:** All controls on a form or report belong to the Controls collection for that Form or Report object. Controls within a particular section belong to the Controls collection for that section. Controls within a tab control or option group control belong to the Controls collection for that control. A label control that is attached to another control belongs to the Controls collection for that control. When you refer to an individual Control object in the Controls collection, you can refer to the Controls collection either implicitly or explicitly. Each Control object is denoted by a particular intrinsic constant. For example, the intrinsic constant acTextBox is associated with a text box control, and acCommandButton is associated with a command button. The constants for the various Microsoft Access controls are set forth in the control's ControlType property. To determine the type of an existing control, you can use the ControlType property. However, you don't need to know the specific type of control to use it in code. You can simply represent it with a variable of data type Control.

## Properties (34)

- `Application As Application  (read-only)`  
  Use the Application property to access the active Microsoft Access Application object and its related properties. Read-only Application object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Column As Variant  (read-only)`  
  Use the Column property to refer to a specific column or column and row combination in a multiple-column combo box or list box. Read-only Variant.
- `Selected As Long  (read/write)`  
  Use the Selected property in Visual Basic to determine if an item is selected. Read/write Long.
- `OldValue As Variant  (read-only)`  
  Use the OldValue property to determine the unedited value of a bound control. Read-only Variant.
- `Form As Form  (read-only)`  
  Use the Form property to refer to a form or to refer to the form associated with a subformcontrol. Read-only Form.
- `Report As Report  (read-only)`  
  Use the Report property to refer to a report or to refer to the report associated with a subreport control. Read-only Report.
- `ItemData As Variant  (read-only)`  
  The ItemData property returns the data in the bound column for the specified row in a combo box or list box. Read-only Variant.
- `Object As Object  (read-only)`  
  Use the Object property in Visual Basic to return a reference to the ActiveX object that is associated with a linked or embedded OLE object in a control. By using this reference, you can access the properties or invoke the methods of the OLE object. Read-only Object.
- `ObjectVerbs As String  (read-only)`  
  Use the ObjectVerbs property in Visual Basic to determine the list of verbs that an OLE object supports. Read-only String.
- `Properties As Properties  (read-only)`  
  Returns a reference to a control's Properties collection object. Read-only.
- `ItemsSelected As _ItemsSelected  (read-only)`  
  Use the ItemsSelected property to return a read-only reference to the hidden ItemsSelected collection. This hidden collection can be used to access data in the selected rows of a multiselect list box control.
- `Pages As Pages  (read-only)`  
  Returns a Pages collection that represents the pages in the specified control that supports tabbed pages (for example, a TabControl object). Read-only.
- `Controls As Children  (read-only)`  
  Returns the Controls collection of a form, subform, report, or section. Read-only Controls.
- `Hyperlink As _Hyperlink  (read-only)`  
  Use the Hyperlink property to return a reference to a Hyperlink object and to access the properties and methods of a control's hyperlink. Read-only.
- `SmartTags As _SmartTags  (read-only)`  
  Returns a SmartTags collection that represents the collection of smart tags that have been added to a control.
- `Layout As AcLayoutType  (read-only)`  
  Returns the type of layout for the specified control. Read-only AcLayoutType.
- `LeftPadding As Integer  (read/write)`  
  Gets or sets the amount of space (in inches) between the control and its left gridline. Read/write Integer.
- `TopPadding As Integer  (read/write)`  
  Gets or sets the amount of space (in inches) between the control and its top gridline. Read/write Integer.
- `RightPadding As Integer  (read/write)`  
  Gets or sets the amount of space (in inches) between the control and its right gridline. Read/write Integer.
- `BottomPadding As Integer  (read/write)`  
  Gets or sets the amount of space (in inches) between the control and its bottom gridline. Read/write Integer.
- `GridlineStyleLeft As Byte  (read/write)`  
  Gets or sets the left gridline style of the specified control. Read/write Byte.
- `GridlineStyleTop As Byte  (read/write)`  
  Gets or sets the top gridline style of the specified control. Read/write Byte.
- `GridlineStyleRight As Byte  (read/write)`  
  Gets or sets the right gridline style of the specified control. Read/write Byte.
- `GridlineStyleBottom As Byte  (read/write)`  
  Gets or sets the bottom gridline style of the specified control. Read/write Byte.
- `GridlineWidthLeft As Byte  (read/write)`  
  Gets or sets the width of the left gridline for the specified control. Read/write Byte.
- `GridlineWidthTop As Byte  (read/write)`  
  Gets or sets the width of the top gridline for the specified control. Read/write Byte.
- `GridlineWidthRight As Byte  (read/write)`  
  Gets or sets the width of the right gridline for the specified control. Read/write Byte.
- `GridlineWidthBottom As Byte  (read/write)`  
  Gets or sets the width of the bottom gridline for the specified control. Read/write Byte.
- `GridlineColor As Long  (read/write)`  
  Gets or sets the color of the gridline for the specified control. Read/write Long.
- `HorizontalAnchor As AcHorizontalAnchor  (read/write)`  
  Gets or sets an AcHorizontalAnchor constant that indicates how the control is anchored horizontally within its layout. Read/write.
- `VerticalAnchor As AcVerticalAnchor  (read/write)`  
  Gets or sets an AcVerticalAnchor constant that indicates how the specified control is anchored vertically within its layout. Read/write.
- `LayoutID As Long  (read-only)`  
  Returns the unique identifier for the layout that contains the specified control. Read-only Long.
- `Name As String  (read/write)`  
  Use the Name property to determine the string expression that identifies the name of an object. Read-only String.

## Methods (6)

- `Undo()`  
  Use the Undo method to reset a control or form when its value has been changed.
- `Dropdown()`  
  Use the Dropdown method to force the list in the specified combo box to drop down.
- `Requery()`  
  The Requery method updates the data underlying a specified control that's on the active form by requerying the source of data for the control.
- `SizeToFit()`  
  Use the SizeToFit method to size a control so that it fits the text or image that it contains.
- `SetFocus()`  
  The SetFocus method moves the focus to the specified form, the specified control on the active form, or the specified field on the active datasheet.
- `Move(Left As Variant, [Top As Variant], [Width As Variant], [Height As Variant])`  
  Moves the specified object to the coordinates specified by the argument values.
    - `Left As Variant` (required): The screen position in twips for the left edge of the object relative to the left edge of the Microsoft Access window.
    - `Top As Variant` (optional): The screen position in twips for the top edge of the object relative to the top edge of the Access window.
    - `Width As Variant` (optional): The desired width of the object in twips.
    - `Height As Variant` (optional): The desired height of the object in twips.
