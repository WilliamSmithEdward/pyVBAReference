# SubForm

**Type:** Class  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {3B06E963-E47C-11CD-8701-00AA003F0F07}  

This object corresponds to a subform control. The subform control embeds a form in a form.

**Remarks:** Instead of creating the main form, and then adding the subform control to it, you can simultaneously create the main form and subform with a wizard. You can also create a subform by dragging an existing form or report from the Database window to the main form.

## Properties (65)

- `Application As Application  (read-only)`  
  Use the Application property to access the active Microsoft Access Application object and its related properties. Read-only Application object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Form As Form  (read-only)`  
  Use the Form property to refer to a form or to refer to the form associated with a subformcontrol. Read-only Form.
- `Report As Report  (read-only)`  
  Use the Report property to refer to a report or to refer to the report associated with a subreport control. Read-only Report.
- `Properties As Properties  (read-only)`  
  Returns a reference to a control's Properties collection object. Read-only.
- `Controls As Controls  (read-only)`  
  Returns the Controls collection of a form, subform, report, or section. Read-only Controls.
- `EventProcPrefix As String  (read/write)`  
  Gets or sets the prefix portion of an event procedure name. Read/write String.
- `ControlType As Byte  (read/write)`  
  Use the ControlType property in Visual Basic to determine the type of control on a form or report. Read/write Byte.
- `SourceObject As String  (read/write)`  
  Use the SourceObject property to identify the form or report that is the source of the subform or subreport on a form or report. Read/write String.
- `LinkChildFields As String  (read/write)`  
  Use the LinkChildFields property (along with the LinkMasterFields property) to specify how Microsoft Access links records in a form or report to records in a subform, subreport, or embedded object, such as a chart. If these properties are set, Access automatically updates the related record in the subform when you change to a new record in a main form. Read/write String.
- `LinkMasterFields As String  (read/write)`  
  Use the LinkMasterFields property (along with the LinkChildFields property) to specify how Microsoft Access links records in a form or report to records in a subform, subreport, or embedded object, such as a chart. If these properties are set, Access automatically updates the related record in the subform when you change to a new record in a main form. Read/write String.
- `StatusBarText As String  (read/write)`  
  Use the StatusBarText property to specify the text that is displayed in the status bar when a control is selected. Read/write String.
- `Visible As Boolean  (read/write)`  
  Returns or sets whether the object is visible. Read/write Boolean.
- `DisplayWhen As Byte  (read/write)`  
  Use the DisplayWhen property to specify which of a form's controls you want displayed on screen and in print. Read/write Byte.
- `Enabled As Boolean  (read/write)`  
  Use the Enabled property to set or return the status of the conditional format in the FormatCondition object. Read/write Boolean.
- `Locked As Boolean  (read/write)`  
  The Locked property specifies whether you can edit data in a control in Form view. Read/write Boolean.
- `TabStop As Boolean  (read/write)`  
  Use the TabStop property to specify whether you can use the Tab key to move the focus to a control. Read/write Boolean.
- `TabIndex As Integer  (read/write)`  
  Use the TabIndex property to specify a control's place in the tab order on a form or report. Read/write Integer.
- `CanGrow As Boolean  (read/write)`  
  Gets or sets whether the specified control automatically adjusts vertically to print or preview all the data that the control contains. Read/write Boolean.
- `CanShrink As Boolean  (read/write)`  
  Gets or sets whether the specified control automatically adjusts vertically to print or preview all the data that the section or control contains. Read/write Boolean.
- `Left As Integer  (read/write)`  
  Use the Left property to specify an object's location on a form or report. Read/write Integer.
- `Top As Integer  (read/write)`  
  Use the Top property to specify an object's location on a form or report. Read/write Integer.
- `Width As Integer  (read/write)`  
  Gets or sets the width of the specified object in twips. Read/write Integer.
- `Height As Integer  (read/write)`  
  Gets or sets the height of the specified object in twips. Read/write Integer.
- `SpecialEffect As Byte  (read/write)`  
  Use the SpecialEffect property to specify whether special formatting will apply to the specified object. Read/write Byte.
- `BorderStyle As Byte  (read/write)`  
  Specifies how a control's border appears. Read/write Byte.
- `OldBorderStyle As Byte  (read/write)`  
  Use this property to set or return the unedited value of the BorderStyle property for a form or control. This property is useful if you need to revert to an unedited or preferred border style. Read/write Byte.
- `BorderWidth As Byte  (read/write)`  
  Use the BorderWidth property to specify the width of a control's border. Read/write Byte.
- `BorderColor As Long  (read/write)`  
  Use the BorderColor property to specify the color of a control's border. Read/write Long.
- `AutoLabel As Boolean  (read/write)`  
  Specifies whether labels are automatically created and attached to new controls. Read/write Boolean.
- `AddColon As Boolean  (read/write)`  
  Specifies whether a colon follows the text in labels for new controls. Read/write Boolean.
- `LabelX As Integer  (read/write)`  
  The LabelX property (along with the LabelY property) specifies the placement of the label for a new control. Read/write Integer.
- `LabelY As Integer  (read/write)`  
  The LabelY property (along with the LabelX property) specifies the placement of the label for a new control. Read/write Integer.
- `LabelAlign As Byte  (read/write)`  
  The LabelAlign property specifies the text alignment within attached labels on new controls. Read/write Byte.
- `Section As Integer  (read/write)`  
  You can identify these controls by the section of a form or report where the control appears. Read/write Integer.
- `Tag As String  (read/write)`  
  Stores extra information about a form, report, section, or control needed by a Microsoft Access application. Read/write String.
- `IsVisible As Boolean  (read/write)`  
  Use the IsVisible property to determine whether a control on a report is visible. Read/write Boolean.
- `InSelection As Boolean  (read/write)`  
  Use the InSelection property to determine or specify whether a control on a form in Design view is selected. Read/write Boolean.
- `OnEnter As String  (read/write)`  
  Sets or returns the value of the On Enter box in the Properties window of specified object. Read/write String.
- `OnExit As String  (read/write)`  
  Sets or returns the value of the On Exit box in the Properties window of specified object. Read/write String.
- `Name As String  (read/write)`  
  Use the Name property to specify or determine the string expression that identifies the name of an object. Read/write String.
- `Layout As AcLayoutType  (read-only)`  
  Returns the type of layout for the specified subform. Read-only AcLayoutType.
- `LeftPadding As Integer  (read/write)`  
  Gets or sets the amount of space (in inches) between the subform and its left gridline. Read/write Integer.
- `TopPadding As Integer  (read/write)`  
  Gets or sets the amount of space (in inches) between the subform and its top gridline. Read/write Integer.
- `RightPadding As Integer  (read/write)`  
  Gets or sets the amount of space (in inches) between the subform and its right gridline. Read/write Integer.
- `BottomPadding As Integer  (read/write)`  
  Gets or sets the amount of space (in inches) between the subform and its bottom gridline. Read/write Integer.
- `GridlineStyleLeft As Byte  (read/write)`  
  Gets or sets the left gridline style of the specified subform. Read/write Byte.
- `GridlineStyleTop As Byte  (read/write)`  
  Gets or sets the top gridline style of the specified subform. Read/write Byte.
- `GridlineStyleRight As Byte  (read/write)`  
  Gets or sets the right gridline style of the specified subform. Read/write Byte.
- `GridlineStyleBottom As Byte  (read/write)`  
  Gets or sets the bottom gridline style of the specified subform. Read/write Byte.
- `GridlineWidthLeft As Byte  (read/write)`  
  Gets or sets the width of the left gridline for the specified subform. Read/write Byte.
- `GridlineWidthTop As Byte  (read/write)`  
  Gets or sets the width of the top gridline for the specified subform. Read/write Byte.
- `GridlineWidthRight As Byte  (read/write)`  
  Gets or sets the width of the right gridline for the specified subform. Read/write Byte.
- `GridlineWidthBottom As Byte  (read/write)`  
  Gets or sets the width of the bottom gridline for the specified subform. Read/write Byte.
- `GridlineColor As Long  (read/write)`  
  Gets or sets the color of the gridline for the specified subform. Read/write Long.
- `HorizontalAnchor As AcHorizontalAnchor  (read/write)`  
  Gets or sets an AcHorizontalAnchor constant that indicates how the subform is anchored horizontally within its layout. Read/write.
- `VerticalAnchor As AcVerticalAnchor  (read/write)`  
  Gets or sets an AcVerticalAnchor constant that indicates how the specified subform is anchored vertically within its layout. Read/write.
- `FilterOnEmptyMaster As Boolean  (read/write)`  
  Gets or sets whether all records are displayed when the subform's master field is Null. Read/write Boolean.
- `LayoutID As Long  (read-only)`  
  Returns the unique identifier for the layout that contains the specified subform. Read-only Long.
- `BorderThemeColorIndex As Long  (read/write)`  
  Gets or sets a value that represents a color in the applied color theme associated with the BorderColor property of the specified object. Read/write Long.
- `BorderTint As Single  (read/write)`  
  Gets or sets the tint that is applied to the theme color in the BorderColor property of the specified object. Read/write Single.
- `BorderShade As Single  (read/write)`  
  Gets or sets the shade that is applied to the theme color in the BorderColor property of the specified object. Read/write Single.
- `GridlineThemeColorIndex As Long  (read/write)`  
  Gets or sets the theme color index that represents a color in the applied color theme associated with the GridlineColor property of the specified object. Read/write Long.
- `GridlineTint As Single  (read/write)`  
  Gets or sets the tint applied to the theme color in the GridlineColor property of the specified object. Read/write Single.
- `GridlineShade As Single  (read/write)`  
  Gets or sets the shade applied to the theme color in the GridlineColor property of the specified object. Read/write Single.

## Methods (4)

- `SizeToFit()`  
  Use the SizeToFit method to size a control so that it fits the text or image that it contains.
- `SetFocus()`  
  The SetFocus method moves the focus to the specified form, the specified control on the active form, or the specified field on the active datasheet.
- `Requery()`  
  The Requery method updates the data underlying a specified subform by requerying the source of data for the subform.
- `Move(Left As Variant, [Top As Variant], [Width As Variant], [Height As Variant])`  
  Moves the specified object to the coordinates specified by the argument values.
    - `Left As Variant` (required): The screen position in twips for the left edge of the object relative to the left edge of the Microsoft Access window.
    - `Top As Variant` (optional): The screen position in twips for the top edge of the object relative to the top edge of the Access window.
    - `Width As Variant` (optional): The desired width of the object in twips.
    - `Height As Variant` (optional): The desired height of the object in twips.

## Events (2)

- `Enter()`  
  The Enter event occurs before a control actually receives the focus from a control on the same form or report.
- `Exit(Cancel As Integer)`  
  The Exit event occurs just before a control loses the focus to another control on the same form or report.
    - `Cancel As Integer` (required): Set to True to cancel the event.
