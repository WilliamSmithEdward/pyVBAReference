# CustomControl

**Type:** Class  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {3B06E967-E47C-11CD-8701-00AA003F0F07}  

When setting the properties of an ActiveX control, you may need or prefer to use the control's custom properties dialog box. This custom properties dialog box provides an alternative to the list of properties in the Microsoft Access property sheet for setting ActiveX control properties in Design view.

**Remarks:** The reason for the custom properties dialog box is that not all applications that use ActiveX controls provide a property sheet like the one in Microsoft Access. The custom properties dialog box provides an interface for setting key control properties regardless of the interface supplied by the hosting application. For some ActiveX control properties, you can choose either of these two locations to set the property: - The Microsoft Access property sheet. - The ActiveX control's custom properties dialog box. In some cases, the custom properties dialog box is the only way to set a property in Design view. This is usually the situation when the interface needed to set a property doesn't work inside the Microsoft Access property sheet. For example, the GridFont property for the Calendar control has a number of arguments; you can't set more than one argument per property in the Microsoft Access property sheet. Not all ActiveX controls provide a custom properties dialog box. To see whether a control provides this custom properties dialog box, look for the Custom property in the Microsoft Access property sheet for this control.

## Properties (68)

- `Application As Application  (read-only)`  
  Use the Application property to access the active Microsoft Access Application object and its related properties. Read-only Application object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `OldValue As Variant  (read-only)`  
  Use the OldValue property to determine the unedited value of a bound control. Read-only Variant.
- `Object As Object  (read-only)`  
  Use the Object property in Visual Basic to return a reference to the ActiveX object that is associated with a linked or embedded OLE object in a control. By using this reference, you can access the properties or invoke the methods of the OLE object. Read-only Object.
- `ObjectVerbs As String  (read-only)`  
  Use the ObjectVerbs property in Visual Basic to determine the list of verbs that an OLE object supports. Read-only String.
- `Properties As Properties  (read-only)`  
  Returns a reference to a control's Properties collection object. Read-only.
- `Controls As Children  (read-only)`  
  Returns the Controls collection of a form, subform, report, or section. Read-only Controls.
- `Value As Variant  (read/write)`  
  Gets or sets the value displayed in the specified control. Read/write Variant.
- `EventProcPrefix As String  (read/write)`  
  Gets or sets the prefix portion of an event procedure name. Read/write String.
- `ControlType As Byte  (read/write)`  
  Use the ControlType property in Visual Basic to determine the type of control on a form or report. Read/write Byte.
- `ControlSource As String  (read/write)`  
  Use the ControlSource property to specify what data appears in a control. You can display and edit data bound to a field in a table, query, or SQL statement. You can also display the result of an expression. Read/write String.
- `OLEClass As String  (read/write)`  
  Use the OLEClass property to obtain a description of the kind of OLE object contained in a chart control or an unbound object frame. Read-only String.
- `Verb As Long  (read/write)`  
  Use the Verb property to specify the operation to perform when an OLE object is activated, which is permitted when the control's Action property is set to acOLEActivate. Read/write Long.
- `Class As String  (read/write)`  
  Use the Class property to specify or determine the class name of an embedded OLE object. Read/write String.
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
- `BorderColor As Long  (read/write)`  
  Use the BorderColor property to specify the color of a control's border. Read/write Long.
- `BorderWidth As Byte  (read/write)`  
  Use the BorderWidth property to specify the width of a control's border. Read/write Byte.
- `ControlTipText As String  (read/write)`  
  Use the ControlTipText property to specify the text that appears in a ScreenTip when you hold the mouse pointer over a control. Read/write String.
- `HelpContextId As Long  (read/write)`  
  The HelpContextID property specifies the context ID of a topic in the custom Help file specified by the HelpFile property setting. Read/write Long.
- `Section As Integer  (read/write)`  
  You can identify these controls by the section of a form or report where the control appears. Read/write Integer.
- `Tag As String  (read/write)`  
  Stores extra information about a form, report, section, or control needed by a Microsoft Access application. Read/write String.
- `ObjectPalette As Variant  (read/write)`  
  The ObjectPalette property specifies the palette in the application used to create an OLE object. Read/write Variant.
- `ObjectVerbsCount As Long  (read/write)`  
  Use the ObjectVerbsCount property in Visual Basic to determine the number of verbs supported by an OLE object. Read-only Long.
- `IsVisible As Boolean  (read/write)`  
  Use the IsVisible property to determine whether a control on a report is visible. Read/write Boolean.
- `InSelection As Boolean  (read/write)`  
  Use the InSelection property to determine or specify whether a control on a form in Design view is selected. Read/write Boolean.
- `OnUpdated As String  (read/write)`  
  Sets or returns the value of the On Updated box in the Properties window of a form or report. Read/write String.
- `OnEnter As String  (read/write)`  
  Sets or returns the value of the On Enter box in the Properties window of specified object. Read/write String.
- `OnExit As String  (read/write)`  
  Sets or returns the value of the On Exit box in the Properties window of specified object. Read/write String.
- `OnGotFocus As String  (read/write)`  
  Sets or returns the value of the On Got Focus box in the Properties window of the specified object. Read/write String.
- `OnLostFocus As String  (read/write)`  
  Sets or returns the value of the On Lost Focus box in the Properties window of the specified object. Read/write String.
- `Default As Boolean  (read/write)`  
  Use the Default property to specify whether a command button is the default button on a form. Read/write Boolean.
- `Cancel As Boolean  (read/write)`  
  Use the Cancel property to specify whether a command button is also the Cancel button on a form. Read/write Boolean.
- `Custom As String  (read/write)`  
  Returns or sets a String representing the custom properties dialog box for an ActiveX control. Read/write.
- `About As String  (read/write)`  
  Returns or sets a String representing version and copyright information for an ActiveX control. Read/write.
- `Name As String  (read/write)`  
  Use the Name property to specify or determine the string expression that identifies the name of an object. Read/write String.
- `BorderThemeColorIndex As Long  (read/write)`  
  Gets or sets a value that represents a color in the applied color theme associated with the BorderColor property of the specified object. Read/write Long.
- `BorderTint As Single  (read/write)`  
  Gets or sets the tint that is applied to the theme color in the BorderColor property of the specified object. Read/write Single.
- `BorderShade As Single  (read/write)`  
  Gets or sets the shade that is applied to the theme color in the BorderColor property of the specified object. Read/write Single.
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
- `VarOleObject As Variant  (read/write)`  
  Gets a pointer to an IOLEObject that represents the memory address of an OLE object. Read-only Variant.

## Methods (4)

- `SizeToFit()`  
  Use the SizeToFit method to size a control so that it fits the text or image that it contains.
- `Requery()`  
  The Requery method updates the data underlying a specified control that's on the active form by requerying the source of data for the control.
- `SetFocus()`  
  The SetFocus method moves the focus to the specified form, the specified control on the active form, or the specified field on the active datasheet.
- `Move(Left As Variant, [Top As Variant], [Width As Variant], [Height As Variant])`  
  Moves the specified object to the coordinates specified by the argument values.
    - `Left As Variant` (required): The screen position in twips for the left edge of the object relative to the left edge of the Microsoft Access window.
    - `Top As Variant` (optional): The screen position in twips for the top edge of the object relative to the top edge of the Access window.
    - `Width As Variant` (optional): The desired width of the object in twips.
    - `Height As Variant` (optional): The desired height of the object in twips.

## Events (5)

- `Updated(Code As Integer)`  
  The Updated event occurs when an OLE object's data has been modified.
- `Enter()`  
  The Enter event occurs before a control actually receives the focus from a control on the same form or report.
- `Exit(Cancel As Integer)`  
  The Exit event occurs just before a control loses the focus to another control on the same form or report.
    - `Cancel As Integer` (required): Set to True to cancel the event.
- `GotFocus()`  
  The GotFocus event occurs when the specified object receives the focus.
- `LostFocus()`  
  The LostFocus event occurs when the specified object loses the focus.
