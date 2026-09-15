# TextBox

**Type:** Class  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {3B06E945-E47C-11CD-8701-00AA003F0F07}  

This object represents a text box control on a form or report. Text boxes are used to display data from a record source, display the results of a calculation, or accept input from a user.

**Remarks:** Text boxes can be either bound or unbound. You use a bound text box to display data from a particular field. You use an unbound text box to display the results of a calculation, or to accept input from a user (as in the following code example).

**Example:**

```vba
Private Sub txtValue1_BeforeUpdate(Cancel As Integer)

MsgBox "The Text box is being updated."

End Sub
```

## Properties (140)

- `Application As Application  (read-only)`  
  Use the Application property to access the active Microsoft Access Application object and its related properties. Read-only Application object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `OldValue As Variant  (read-only)`  
  Use the OldValue property to determine the unedited value of a bound control. Read-only Variant.
- `Properties As Properties  (read-only)`  
  Returns a reference to a control's Properties collection object. Read-only.
- `Controls As Children  (read-only)`  
  Returns the Controls collection of a form, subform, report, or section. Read-only Controls.
- `Hyperlink As _Hyperlink  (read-only)`  
  Use the Hyperlink property to return a reference to a Hyperlink object and to access the properties and methods of a control's hyperlink. Read-only.
- `FormatConditions As FormatConditions  (read-only)`  
  Use the FormatConditions property to return a read-only reference to the FormatConditions collection and its related properties.
- `Value As Variant  (read/write)`  
  Determines or specifies the text in the text box. Read/write Variant.
- `EventProcPrefix As String  (read/write)`  
  Gets or sets the prefix portion of an event procedure name. Read/write String.
- `ControlType As Byte  (read/write)`  
  Use the ControlType property in Visual Basic to determine the type of control on a form or report. Read/write Byte.
- `ControlSource As String  (read/write)`  
  Use the ControlSource property to specify what data appears in a control. You can display and edit data bound to a field in a table, query, or SQL statement. You can also display the result of an expression. Read/write String.
- `Format As String  (read/write)`  
  Use the Format property to customize the way numbers, dates, times, and text are displayed and printed. Read/write String.
- `DecimalPlaces As Byte  (read/write)`  
  Use the DecimalPlaces property to specify the number of decimal places that Microsoft Access uses to display numbers. Read/write Byte.
- `InputMask As String  (read/write)`  
  Use the InputMask property to make data entry easier and to control the values that users can enter in a text box control. Read/write String.
- `DefaultValue As String  (read/write)`  
  Specifies a value that is automatically entered in a field when a new record is created. For example, in an Addresses table, you can set the default value for the City field to New York. When users add a record to the table, they can either accept this value or enter the name of a different city. Read/write String.
- `IMEHold As Boolean  (read/write)`  
  Use the IMEHold/Hold KanjiConversionMode property to show whether the Kanji Conversion Mode is maintained when the control loses the focus. Read/write Boolean.
- `FuriganaControl As String  (read/write)`
- `PostalAddress As String  (read/write)`  
  Use the PostalAddress property to specify or determine the postal code and the Customer Barcode data corresponding to the address information displayed in a specified field or text box. The PostalAddress Property Wizard enables the setting of these properties. Read/write String.
- `ValidationRule As String  (read/write)`  
  Use the ValidationRule property to specify requirements for data entered into a record, field, or control. When data is entered that violates the ValidationRule setting, you can use the ValidationText property to specify the message to be displayed to the user. Read/write String.
- `ValidationText As String  (read/write)`  
  Use the ValidationText property to specify a message to be displayed to the user when data is entered that violates a ValidationRule setting for a record, field, or control. Read/write String.
- `StatusBarText As String  (read/write)`  
  Use the StatusBarText property to specify the text that is displayed in the status bar when a control is selected. Read/write String.
- `EnterKeyBehavior As Boolean  (read/write)`  
  Use the EnterKeyBehavior property to specify what happens when you press Enter in a text box control in Form view or Datasheet view. Read/write Boolean.
- `AllowAutoCorrect As Boolean  (read/write)`  
  Use the AllowAutoCorrect property to specify whether the specified control will automatically correct entries made by the user. Read/write Boolean.
- `Visible As Boolean  (read/write)`  
  Returns or sets whether the object is visible. Read/write Boolean.
- `DisplayWhen As Byte  (read/write)`  
  Use the DisplayWhen property to specify which of a form's controls you want displayed on screen and in print. Read/write Byte.
- `Vertical As Boolean  (read/write)`  
  Use the Vertical property to set a form control for vertical display and editing, or to set a report control for vertical display and printing. Read/write Boolean.
- `Enabled As Boolean  (read/write)`  
  Use the Enabled property to set or return the status of the conditional format in the FormatCondition object. Read/write Boolean.
- `Locked As Boolean  (read/write)`  
  The Locked property specifies whether you can edit data in a control in Form view. Read/write Boolean.
- `FilterLookup As Byte  (read/write)`  
  Use the FilterLookup property to specify whether values appear in a bound text box control when using the Filter By Form or Server Filter By Form window. Read/write Byte.
- `AutoTab As Boolean  (read/write)`  
  Use the AutoTab property to specify whether an automatic tab occurs when the last character permitted by a text box control's input mask is entered. An automatic tab moves the focus to the next control in the form's tab order. Read/write Boolean.
- `TabStop As Boolean  (read/write)`  
  Use the TabStop property to specify whether you can use the Tab key to move the focus to a control. Read/write Boolean.
- `TabIndex As Integer  (read/write)`  
  Use the TabIndex property to specify a control's place in the tab order on a form or report. Read/write Integer.
- `ScrollBars As Byte  (read/write)`  
  Use the ScrollBars property to specify whether scroll bars appear on a text box control. Read/write Byte.
- `HideDuplicates As Boolean  (read/write)`  
  Use the HideDuplicates property to hide a control on a report when its value is the same as in the preceding record. Read/write Boolean.
- `CanGrow As Boolean  (read/write)`  
  Gets or sets whether the specified control automatically adjusts vertically to print or preview all the data that the control contains. Read/write Boolean.
- `CanShrink As Boolean  (read/write)`  
  Gets or sets whether the specified control automatically adjusts vertically to print or preview all the data that the section or control contains. Read/write Boolean.
- `RunningSum As Byte  (read/write)`  
  Use the RunningSum property to calculate record-by-record or group-by-group totals in a report. Read/write Byte.
- `Left As Integer  (read/write)`  
  Use the Left property to specify an object's location on a form or report. Read/write Integer.
- `Top As Integer  (read/write)`  
  Use the Top property to specify an object's location on a form or report. Read/write Integer.
- `Width As Integer  (read/write)`  
  Gets or sets the width of the specified object in twips. Read/write Integer.
- `Height As Integer  (read/write)`  
  Gets or sets the height of the specified object in twips. Read/write Integer.
- `BackStyle As Byte  (read/write)`  
  Use the BackStyle property to specify whether a control will be transparent. Read/write Byte.
- `BackColor As Long  (read/write)`  
  Gets or sets the interior color of the specified object. Read/write Long.
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
- `ForeColor As Long  (read/write)`  
  Use the ForeColor property to specify the color for text in a control. Read/write Long.
- `FontName As String  (read/write)`  
  Use the FontName property to specify the font for text in the following situations:
- `FontSize As Integer  (read/write)`  
  Use the FontSize property to specify the point size for text in the following situations:
- `FontWeight As Integer  (read/write)`  
  Use the FontWeight property to specify the line width that Windows uses to display and print characters in a control. Read/write Integer.
- `FontItalic As Boolean  (read/write)`  
  Use the FontItalic property to specify whether text is italic in the following situations:
- `FontUnderline As Boolean  (read/write)`  
  Use the FontUnderline property to specify whether text is underlined in the following situations:
- `TextAlign As Byte  (read/write)`  
  The TextAlign property specifies the text alignment in new controls. Read/write Byte.
- `FontBold As Integer  (read/write)`  
  Use the FontBold property to specify whether a font appears in a bold style in the following situations:
- `ShortcutMenuBar As String  (read/write)`  
  Use the ShortcutMenuBar property to specify the shortcut menu that appears when you right-click the specified object. Read/write String.
- `ControlTipText As String  (read/write)`  
  Use the ControlTipText property to specify the text that appears in a ScreenTip when you hold the mouse pointer over a control. Read/write String.
- `HelpContextId As Long  (read/write)`  
  The HelpContextID property specifies the context ID of a topic in the custom Help file specified by the HelpFile property setting. Read/write Long.
- `ColumnWidth As Integer  (read/write)`  
  Use the ColumnWidth property to specify the width of a column in Datasheet view. Read/write Integer.
- `ColumnOrder As Integer  (read/write)`  
  Use the ColumnOrder property to specify the order of the columns in Datasheet view. Read/write Integer.
- `ColumnHidden As Boolean  (read/write)`  
  Use the ColumnHidden property to show or hide a specified column in Datasheet view. Read/write Boolean.
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
- `Text As String  (read/write)`  
  Use the Text property to set or return the text Value contained in a text box. Read/write String.
- `SelText As String  (read/write)`  
  The SelText property returns a string containing the selected text. Read/write String.
- `SelStart As Integer  (read/write)`  
  The SelStart property specifies or determines the starting point of the selected text or the position of the insertion point if no text is selected. Read/write Integer.
- `SelLength As Integer  (read/write)`  
  The SelLength property specifies or determines the number of characters selected in a text box. Read/write Integer.
- `IsVisible As Boolean  (read/write)`  
  Use the IsVisible property to determine whether a control on a report is visible. Read/write Boolean.
- `InSelection As Boolean  (read/write)`  
  Use the InSelection property to determine or specify whether a control on a form in Design view is selected. Read/write Boolean.
- `BeforeUpdate As String  (read/write)`
- `AfterUpdate As String  (read/write)`
- `OnChange As String  (read/write)`  
  Sets or returns the value of the On Change box in the Properties window of one of the objects in the Applies To list. Read/write String.
- `OnEnter As String  (read/write)`  
  Sets or returns the value of the On Enter box in the Properties window of specified object. Read/write String.
- `OnExit As String  (read/write)`  
  Sets or returns the value of the On Exit box in the Properties window of specified object. Read/write String.
- `OnGotFocus As String  (read/write)`  
  Sets or returns the value of the On Got Focus box in the Properties window of the specified object. Read/write String.
- `OnLostFocus As String  (read/write)`  
  Sets or returns the value of the On Lost Focus box in the Properties window of the specified object. Read/write String.
- `OnClick As String  (read/write)`  
  Sets or returns the value of the On Click box in the Properties window. Read/write String.
- `OnDblClick As String  (read/write)`  
  Sets or returns the value of the On Dbl Click box in the Properties window. Read/write String.
- `OnMouseDown As String  (read/write)`  
  Sets or returns the value of the On Mouse Down box in the Properties window. Read/write String.
- `OnMouseMove As String  (read/write)`  
  Sets or returns the value of the On Mouse Move box in the Properties window. Read/write String.
- `OnMouseUp As String  (read/write)`  
  Sets or returns the value of the On Mouse Up box in the Properties window. Read/write String.
- `OnKeyDown As String  (read/write)`  
  Sets or returns the value of the On Key Down box in the Properties window. Read/write String.
- `OnKeyUp As String  (read/write)`  
  Sets or returns the value of the On Key Up box in the Properties window. Read/write String.
- `OnKeyPress As String  (read/write)`  
  Sets or returns the value of the On Key Press box in the Properties window. Read/write String.
- `ReadingOrder As Byte  (read/write)`  
  Use the ReadingOrder property to specify or determine the reading order of words in text. Read/write Byte.
- `KeyboardLanguage As Byte  (read/write)`
- `ScrollBarAlign As Byte  (read/write)`  
  Use the ScrollBarAlign property to specify or determine the alignment of a vertical scroll bar. Read/write Byte.
- `NumeralShapes As Byte  (read/write)`
- `IMEMode As AcImeMode  (read/write)`
- `Name As String  (read/write)`  
  Use the Name property to specify or determine the string expression that identifies the name of an object. Read/write String.
- `IMESentenceMode As AcImeSentenceMode  (read/write)`
- `LeftMargin As Integer  (read/write)`  
  Along with the TopMargin, RightMargin, and BottomMargin properties, specifies the location of information displayed within a text box control. Read/write Integer.
- `TopMargin As Integer  (read/write)`  
  Along with the LeftMargin, RightMargin, and BottomMargin properties, specifies the location of information displayed within a text box control. Read/write Integer.
- `LineSpacing As Integer  (read/write)`  
  Use the LineSpacing property to specify or determine the location of information displayed within a label or text box control. Read/write Integer.
- `RightMargin As Integer  (read/write)`  
  Along with the TopMargin, LeftMargin, and BottomMargin properties, specifies the location of information displayed within a text box control. Read/write Integer.
- `BottomMargin As Integer  (read/write)`  
  Along with the LeftMargin, RightMargin, and TopMargin properties, specifies the location of information displayed within a text box control. Read/write Integer.
- `IsHyperlink As Boolean  (read/write)`  
  Use the IsHyperlink property to specify or determine if the data contained in a text box is a hyperlink. Read/write Boolean.
- `OnDirty As String  (read/write)`  
  Sets or returns the value of the On Dirty box in the Properties window of a form or report. Read/write String.
- `OnUndo As String  (read/write)`  
  Returns or sets a String indicating which macro, event procedure, or user-defined function runs when the Undo event occurs. Read/write.
- `AsianLineBreak As Boolean  (read/write)`  
  Returns or sets a Boolean indicating whether line breaks in text boxes follow rules governing East Asian languages. True to control line breaks based on East Asian language rules. Read/write.
- `SmartTags As _SmartTags  (read-only)`  
  Returns a SmartTags collection that represents the collection of smart tags that have been added to a control.
- `Layout As AcLayoutType  (read-only)`  
  Returns the type of layout for the specified text box. Read-only AcLayoutType.
- `TextFormat As AcTextFormat  (read/write)`  
  Gets or sets whether rich text is displayed in the specified text box. Read/write AcTextFormat.
- `LeftPadding As Integer  (read/write)`  
  Gets or sets the amount of space (in inches) between the text box and its left gridline. Read/write Integer.
- `TopPadding As Integer  (read/write)`  
  Gets or sets the amount of space (in inches) between the text box and its top gridline. Read/write Integer.
- `RightPadding As Integer  (read/write)`  
  Gets or sets the amount of space (in inches) between the text box and its right gridline. Read/write Integer.
- `BottomPadding As Integer  (read/write)`  
  Gets or sets the amount of space (in inches) between the text box and its bottom gridline. Read/write Integer.
- `GridlineStyleLeft As Byte  (read/write)`  
  Gets or sets the left gridline style of the specified text box. Read/write Byte.
- `GridlineStyleTop As Byte  (read/write)`  
  Gets or sets the top gridline style of the specified text box. Read/write Byte.
- `GridlineStyleRight As Byte  (read/write)`  
  Gets or sets the right gridline style of the specified text box. Read/write Byte.
- `GridlineStyleBottom As Byte  (read/write)`  
  Gets or sets the bottom gridline style of the specified text box. Read/write Byte.
- `GridlineWidthLeft As Byte  (read/write)`  
  Gets or sets the width of the left gridline for the specified text box. Read/write Byte.
- `GridlineWidthTop As Byte  (read/write)`  
  Gets or sets the width of the top gridline for the specified text box. Read/write Byte.
- `GridlineWidthRight As Byte  (read/write)`  
  Gets or sets the width of the right gridline for the specified text box. Read/write Byte.
- `GridlineWidthBottom As Byte  (read/write)`  
  Gets or sets the width of the bottom gridline for the specified text box. Read/write Byte.
- `GridlineColor As Long  (read/write)`  
  Gets or sets the color of the gridline for the specified text box. Read/write Long.
- `HorizontalAnchor As AcHorizontalAnchor  (read/write)`  
  Gets or sets an AcHorizontalAnchor constant that indicates how the text box is anchored horizontally within its layout. Read/write.
- `VerticalAnchor As AcVerticalAnchor  (read/write)`  
  Gets or sets an AcVerticalAnchor constant that indicates how the specified text box is anchored vertically within its layout. Read/write.
- `LayoutID As Long  (read-only)`  
  Returns the unique identifier for the layout that contains the specified text box. Read-only Long.
- `DisplayAsHyperlink As AcDisplayAsHyperlink  (read/write)`  
  Gets or sets an AcDisplayAsHyperlink constant that specifies whether to display the contents of the specified text box as a hyperlink. Read/write.
- `ShowDatePicker As Integer  (read/write)`  
  Gets or sets whether the date picker control is displayed for the specified text box. Read/write Integer.
- `BackThemeColorIndex As Long  (read/write)`  
  Gets or sets a value that represents a color in the applied color theme associated with the BackColor property of the specified object. Read/write Long.
- `BackTint As Single  (read/write)`  
  Gets or sets the tint that is applied to the theme color in the BackColor property of the specified object. Read/write Single.
- `BackShade As Single  (read/write)`  
  Gets or sets the shade applied to the theme color in the BackColor property of the specified object. Read/write Single.
- `BorderThemeColorIndex As Long  (read/write)`  
  Gets or sets a value that represents a color in the applied color theme associated with the BorderColor property of the specified object. Read/write Long.
- `BorderTint As Single  (read/write)`  
  Gets or sets the tint that is applied to the theme color in the BorderColor property of the specified object. Read/write Single.
- `BorderShade As Single  (read/write)`  
  Gets or sets the shade that is applied to the theme color in the BorderColor property of the specified object. Read/write Single.
- `ForeThemeColorIndex As Long  (read/write)`  
  Gets or sets a value that represents a color in the applied color theme associated with the ForeColor property of the specified object. Read/write Long.
- `ForeTint As Single  (read/write)`  
  Gets or sets the tint that is applied to the theme color in the ForeColor property of the specified object. Read/write Single.
- `ForeShade As Single  (read/write)`  
  Gets or sets the shade that is applied to the theme color in the ForeColor property of the specified object. Read/write Single.
- `ThemeFontIndex As Long  (read/write)`  
  Gets or sets the font index that represents a font in the applied theme associated with the FontName property of the specified object. Read/write Long.
- `GridlineThemeColorIndex As Long  (read/write)`  
  Gets or sets the theme color index that represents a color in the applied color theme associated with the GridlineColor property of the specified object. Read/write Long.
- `GridlineTint As Single  (read/write)`  
  Gets or sets the tint applied to the theme color in the GridlineColor property of the specified object. Read/write Single.
- `GridlineShade As Single  (read/write)`  
  Gets or sets the shade applied to the theme color in the GridlineColor property of the specified object. Read/write Single.

## Methods (5)

- `Undo()`  
  Use the Undo method to reset a control or form when its value has been changed.
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

## Events (17)

- `BeforeUpdate(Cancel As Integer)`
- `AfterUpdate()`
- `Change()`  
  The Change event occurs when the contents of the specified control change.
- `Enter()`  
  The Enter event occurs before a control actually receives the focus from a control on the same form or report.
- `Exit(Cancel As Integer)`  
  The Exit event occurs just before a control loses the focus to another control on the same form or report.
    - `Cancel As Integer` (required): Set to True to cancel the event.
- `GotFocus()`  
  The GotFocus event occurs when the specified object receives the focus.
- `LostFocus()`  
  The LostFocus event occurs when the specified object loses the focus.
- `Click()`  
  The Click event occurs when the user presses and then releases a mouse button over an object.
- `DblClick(Cancel As Integer)`  
  The DblClick event occurs when the user presses and releases the left mouse button twice over an object within the double-click time limit of the system.
    - `Cancel As Integer` (required): The setting determines if the DblClick event occurs. Setting the _Cancel_ argument to True (1) cancels the DblClick event.
- `MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)`  
  The MouseDown event occurs when the user presses a mouse button.
    - `Button As Integer` (required): The button that was pressed to trigger the event. If you need to test for the _Button_ argument, you can use one of the following intrinsic constants as bit masks:<ul><li><p><b>acLeftButton</b> The bit mask for the left mouse button.</p></li><li><p><b>acRightButton</b> The bit mask for the right mouse button.</p></li><li><p><b>acMiddleButton</b> The bit mask for the middle mouse button.</p></li></ul>
    - `Shift As Integer` (required): The state of the Shift, Ctrl, and Alt keys when the button specified by the _Button_ argument was pressed or released. If you need to test for the _Shift_ argument, you can use one of the following intrinsic constants as bit masks:<ul><li><p><b>acShiftMask</b> The bit mask for the Shift key.</p></li><li><p><b>acCtrlMask</b> The bit mask for the Ctrl key.</p></li><li><p><b>acAltMask</b> The bit mask for the Alt key.</p></li></ul>
    - `X As Single` (required): The _x_ coordinate for the current location of the mouse pointer, in twips.
    - `Y As Single` (required): The _y_ coordinate for the current location of the mouse pointer, in twips.
- `MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)`  
  The MouseMove event occurs when the user moves the mouse.
    - `Button As Integer` (required): The button that was pressed or released when the event was triggered. If you need to test for the _Button_ argument, you can use one of the following intrinsic constants as bit masks:<ul><li><p><b>acLeftButton</b> The bit mask for the left mouse button.</p></li><li><p><b>acRightButton</b> The bit mask for the right mouse button.</p></li><li><p><b>acMiddleButton</b> The bit mask for the middle mouse button.</p></li></ul>
    - `Shift As Integer` (required): The state of the Shift, Ctrl, and Alt keys when the button specified by the _Button_ argument was pressed or released. If you need to test for the _Shift_ argument, you can use one of the following intrinsic constants as bit masks:<ul><li><p><b>acShiftMask</b> The bit mask for the Shift key.</p></li><li><p><b>acCtrlMask</b> The bit mask for the Ctrl key.</p></li><li><p><b>acAltMask</b> The bit mask for the Alt key.</p></li></ul>
    - `X As Single` (required): The _x_ coordinate for the current location of the mouse pointer, in twips.
    - `Y As Single` (required): The _y_ coordinate for the current location of the mouse pointer, in twips.
- `MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)`  
  The MouseUp event occurs when the user releases a mouse button.
    - `Button As Integer` (required): The button that was released to trigger the event. If you need to test for the _Button_ argument, you can use one of the following intrinsic constants as bit masks:<ul><li><p><b>acLeftButton</b> The bit mask for the left mouse button.</p></li><li><p><b>acRightButton</b> The bit mask for the right mouse button.</p></li><li><p><b>acMiddleButton</b> The bit mask for the middle mouse button.</p></li></ul>
    - `Shift As Integer` (required): The state of the Shift, Ctrl, and Alt keys when the button specified by the _Button_ argument was pressed or released. If you need to test for the _Shift_ argument, you can use one of the following intrinsic constants as bit masks:<ul><li><p><b>acShiftMask</b> The bit mask for the Shift key.</p></li><li><p><b>acCtrlMask</b> The bit mask for the Ctrl key.</p></li><li><p><b>acAltMask</b> The bit mask for the Alt key.</p></li></ul>
    - `X As Single` (required): The _x_ coordinate for the current location of the mouse pointer, in twips.
    - `Y As Single` (required): The _y_ coordinate for the current location of the mouse pointer, in twips.
- `KeyDown(KeyCode As Integer, Shift As Integer)`  
  The KeyDown event occurs when the user presses a key while a form or control has the focus. This event also occurs if you send a keystroke to a form or control by using the SendKeys action in a macro or the SendKeys statement in Visual Basic.
    - `KeyCode As Integer` (required): A key code, such as vbKeyF1 (the F1 key) or vbKeyHome (the Home key). To specify key codes, use the intrinsic constants shown in the Object Browser. You can prevent an object from receiving a keystroke by setting _KeyCode_ to 0.
    - `Shift As Integer` (required): The state of the Shift, Ctrl, and Alt keys at the time of the event. If you need to test for the _Shift_ argument, you can use one of the following intrinsic constants as bit masks:<ul><li><p>acShiftMask The bit mask for the Shift key.</p></li><li><p>acCtrlMask The bit mask for the Ctrl key.</p></li><li><p>acAltMask The bit mask for the Alt key.</p></li></ul>
- `KeyPress(KeyAscii As Integer)`  
  The KeyPress event occurs when the user presses and releases a key or key combination that corresponds to an ANSI code while a form or control has the focus. This event also occurs if you send an ANSI keystroke to a form or control by using the SendKeys action in a macro or the SendKeys statement in Visual Basic.
    - `KeyAscii As Integer` (required): Returns a numeric ANSI key code. The _KeyAscii_ argument is passed by reference; changing it sends a different character to the object. Setting the _KeyAscii_ argument to 0 cancels the keystroke so that the object doesn't recognize that a key was pressed.
- `KeyUp(KeyCode As Integer, Shift As Integer)`  
  The KeyUp event occurs when the user releases a key while a form or control has the focus. This event also occurs if you send a keystroke to a form or control by using the SendKeys action in a macro or the SendKeys statement in Visual Basic.
    - `KeyCode As Integer` (required): A key code, such as vbKeyF1 (the F1 key) or vbKeyHome (the Home key). To specify key codes, use the intrinsic constants shown in the Object Browser. You can prevent an object from receiving a keystroke by setting _KeyCode_ to 0.
    - `Shift As Integer` (required): The state of the Shift, Ctrl, and Alt keys at the time of the event. If you need to test for the _Shift_ argument, you can use one of the following intrinsic constants as bit masks:<ul><li><p><b>acShiftMask</b> The bit mask for the Shift key.</p></li><li><p><b>acCtrlMask</b> The bit mask for the Ctrl key.</p></li><li><p><b>acAltMask</b> The bit mask for the Alt key.</p></li></ul>
- `Dirty(Cancel As Integer)`  
  The Dirty event occurs when the contents of the specified control changes.
    - `Cancel As Integer` (required): The setting determines if the Dirty event occurs. Setting the _Cancel_ argument to True (1) cancels the Dirty event.
- `Undo(Cancel As Integer)`  
  Occurs when the user undoes a change.
    - `Cancel As Integer` (required): Set this argument to True to cancel the undo operation and leave the control or form in its edited state.
