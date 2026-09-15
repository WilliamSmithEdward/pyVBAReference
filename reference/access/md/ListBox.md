# ListBox

**Type:** Class  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {3B06E959-E47C-11CD-8701-00AA003F0F07}  

This object corresponds to a list box control. The list box control displays a list of values or alternatives.

**Remarks:** In many cases, it's quicker and easier to select a value from a list than to remember a value to type. A list of choices also helps ensure that the value that's entered in a field is correct. The list in a list box consists of rows of data. Rows can have one or more columns, which can appear with or without headings, as shown in the following diagram. !Multi-column list box If a multiple-column list box is bound, Microsoft Access stores the values from one of the columns. Use an unbound list box to store a value that you can use with another control. For example, you could use an unbound list box to limit the values in another list box or in a custom dialog box. You could also use an unbound list box to find a record based on the value that you select in the list box. If you don't have room on your form to display a list box, or if you want to be able to type new values as well as select values from a list, use a combo box instead of a list box.

**Example:**

```vba
Private Sub FilterBy_Change()

    Dim sql As String

    'This will match any entry in the list that begins with what the user
    'has typed in the FilterBy control
    sql = "SELECT ColorID, ColorName FROM Colors WHERE ColorName Like '" & Me.FilterBy.Text & "*' ORDER BY ColorName"

    'If you want to match any part of the string then add wildcard (*) before
    'the FilterBy.Text, too:
    'sql = "SELECT ColorID, ColorName FROM Colors WHERE ColorName Like '*" & Me.FilterBy.Text & "*' ORDER BY ColorName"

    Me.ColorID.RowSource = sql

End Sub
```

## Properties (123)

- `Application As Application  (read-only)`  
  Use the Application property to access the active Microsoft Access Application object and its related properties. Read-only Application object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Column As Variant  (read-only)`  
  Use the Column property to refer to a specific column or column and row combination in a multiple-column combo box or list box. Read-only Variant.
- `Selected As Long  (read/write)`  
  Use the Selected property in Visual Basic to determine if an item in a list box is selected. Read/write Long.
- `OldValue As Variant  (read-only)`  
  Use the OldValue property to determine the unedited value of a bound control. Read-only Variant.
- `ItemData As Variant  (read-only)`  
  The ItemData property returns the data in the bound column for the specified row in a list box. Read-only Variant.
- `Properties As Properties  (read-only)`  
  Returns a reference to a control's Properties collection object. Read-only.
- `Controls As Children  (read-only)`  
  Returns the Controls collection of a form, subform, report, or section. Read-only Controls.
- `ItemsSelected As _ItemsSelected  (read-only)`  
  Use the ItemsSelected property to return a read-only reference to the hidden ItemsSelected collection. This hidden collection can be used to access data in the selected rows of a multiselect list box control.
- `Hyperlink As _Hyperlink  (read-only)`  
  Use the Hyperlink property to return a reference to a Hyperlink object and to access the properties and methods of a control's hyperlink. Read-only.
- `Value As Variant  (read/write)`  
  Determines or specifies which value or option in the list box is selected. Read/write Variant.
- `EventProcPrefix As String  (read/write)`  
  Gets or sets the prefix portion of an event procedure name. Read/write String.
- `ControlType As Byte  (read/write)`  
  Use the ControlType property in Visual Basic to determine the type of control on a form or report. Read/write Byte.
- `ControlSource As String  (read/write)`  
  Use the ControlSource property to specify what data appears in a control. You can display and edit data bound to a field in a table, query, or SQL statement. You can also display the result of an expression. Read/write String.
- `RowSourceType As String  (read/write)`  
  Use the RowSourceType property (along with the RowSource property) to tell Microsoft Access how to provide data to the specified object. Read/write String.
- `RowSource As String  (read/write)`  
  Use the RowSource property (along with the RowSourceType property) to tell Microsoft Access how to provide data to the specified object. Read/write String.
- `ColumnCount As Integer  (read/write)`  
  Use the ColumnCount property to specify the number of columns displayed in a list box or in the list box portion of a combo box, or sent to OLE objects in a chart control or unbound object frame. Read/write Integer.
- `ColumnHeads As Boolean  (read/write)`  
  Use the ColumnHeads property to display a single row of column headings for list boxes, combo boxes, and OLE objects that accept column headings. You can also use this property to create a label for each entry in a chart control. What is actually displayed as the first-row column heading depends on the object's RowSourceType property setting. Read/write Boolean.
- `ColumnWidths As String  (read/write)`  
  Use the ColumnWidths property to specify the width of each column in a multiple-column list box. Read/write String.
- `BoundColumn As Long  (read/write)`  
  When you make a selection from a list box, the BoundColumn property tells Microsoft Access which column's values to use as the value of the control. If the control is bound to a field, the value in the column specified by the BoundColumn property is stored in the field named in the ControlSource property. Read/write Long.
- `DefaultValue As String  (read/write)`  
  Specifies a value that is automatically entered in a field when a new record is created. For example, in an Addresses table, you can set the default value for the City field to New York. When users add a record to the table, they can either accept this value or enter the name of a different city. Read/write String.
- `IMEHold As Boolean  (read/write)`  
  Use the IMEHold/Hold KanjiConversionMode property to show whether the Kanji Conversion Mode is maintained when the control loses the focus. Read/write Boolean.
- `ValidationRule As String  (read/write)`  
  Use the ValidationRule property to specify requirements for data entered into a record, field, or control. When data is entered that violates the ValidationRule setting, you can use the ValidationText property to specify the message to be displayed to the user. Read/write String.
- `ValidationText As String  (read/write)`  
  Use the ValidationText property to specify a message to be displayed to the user when data is entered that violates a ValidationRule setting for a record, field, or control. Read/write String.
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
- `MultiSelect As Byte  (read/write)`  
  Use the MultiSelect property to specify whether a user can make multiple selections in a list box on a form and how the multiple selections can be made. Read/write Byte.
- `TabStop As Boolean  (read/write)`  
  Use the TabStop property to specify whether you can use the Tab key to move the focus to a control. Read/write Boolean.
- `TabIndex As Integer  (read/write)`  
  Use the TabIndex property to specify a control's place in the tab order on a form or report. Read/write Integer.
- `HideDuplicates As Boolean  (read/write)`  
  Use the HideDuplicates property to hide a control on a report when its value is the same as in the preceding record. Read/write Boolean.
- `Left As Integer  (read/write)`  
  Use the Left property to specify an object's location on a form or report. Read/write Integer.
- `Top As Integer  (read/write)`  
  Use the Top property to specify an object's location on a form or report. Read/write Integer.
- `Width As Integer  (read/write)`  
  Gets or sets the width of the specified object in twips. Read/write Integer.
- `Height As Integer  (read/write)`  
  Gets or sets the height of the specified object in twips. Read/write Integer.
- `BackColor As Long  (read/write)`  
  Gets or sets the interior color of the specified object. Read/write Long.
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
- `ListCount As Long  (read/write)`  
  Use the ListCount property to determine the number of rows in a list box. Read/write Long.
- `ListIndex As Long  (read/write)`  
  Use the ListIndex property to determine which item is selected in a list box. Read/write Long.
- `IsVisible As Boolean  (read/write)`  
  Use the IsVisible property to determine whether a control on a report is visible. Read/write Boolean.
- `InSelection As Boolean  (read/write)`  
  Use the InSelection property to determine or specify whether a control on a form in Design view is selected. Read/write Boolean.
- `BeforeUpdate As String  (read/write)`
- `AfterUpdate As String  (read/write)`
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
- `ScrollBarAlign As Byte  (read/write)`  
  Use the ScrollBarAlign property to specify or determine the alignment of a vertical scroll bar. Read/write Byte.
- `NumeralShapes As Byte  (read/write)`
- `IMEMode As AcImeMode  (read/write)`
- `Name As String  (read/write)`  
  Use the Name property to specify or determine the string expression that identifies the name of an object. Read/write String.
- `IMESentenceMode As AcImeSentenceMode  (read/write)`
- `Recordset As Object  (read/write)`  
  Returns or sets the ADO Recordset or DAO Recordset object that represents the record source for the specified object. Read/write Object.
- `SmartTags As _SmartTags  (read-only)`  
  Returns a SmartTags collection that represents the collection of smart tags that have been added to a control.
- `Layout As AcLayoutType  (read-only)`  
  Returns the type of layout for the specified list box. Read-only AcLayoutType.
- `LeftPadding As Integer  (read/write)`  
  Gets or sets the amount of space (in inches) between the list box and its left gridline. Read/write Integer.
- `TopPadding As Integer  (read/write)`  
  Gets or sets the amount of space (in inches) between the list box and its top gridline. Read/write Integer.
- `RightPadding As Integer  (read/write)`  
  Gets or sets the amount of space (in inches) between the list box and its right gridline. Read/write Integer.
- `BottomPadding As Integer  (read/write)`  
  Gets or sets the amount of space (in inches) between the list box and its bottom gridline. Read/write Integer.
- `GridlineStyleLeft As Byte  (read/write)`  
  Gets or sets the left gridline style of the specified list box. Read/write Byte.
- `GridlineStyleTop As Byte  (read/write)`  
  Gets or sets the top gridline style of the specified list box. Read/write Byte.
- `GridlineStyleRight As Byte  (read/write)`  
  Gets or sets the right gridline style of the specified list box. Read/write Byte.
- `GridlineStyleBottom As Byte  (read/write)`  
  Gets or sets the bottom gridline style of the specified list box. Read/write Byte.
- `GridlineWidthLeft As Byte  (read/write)`  
  Gets or sets the width of the left gridline for the specified list box. Read/write Byte.
- `GridlineWidthTop As Byte  (read/write)`  
  Gets or sets the width of the top gridline for the specified list box. Read/write Byte.
- `GridlineWidthRight As Byte  (read/write)`  
  Gets or sets the width of the right gridline for the specified list box. Read/write Byte.
- `GridlineWidthBottom As Byte  (read/write)`  
  Gets or sets the width of the bottom gridline for the specified list box. Read/write Byte.
- `GridlineColor As Long  (read/write)`  
  Gets or sets the color of the gridline for the specified list box. Read/write Long.
- `HorizontalAnchor As AcHorizontalAnchor  (read/write)`  
  Gets or sets an AcHorizontalAnchor constant that indicates how the list box is anchored horizontally within its layout. Read/write.
- `VerticalAnchor As AcVerticalAnchor  (read/write)`  
  Gets or sets an AcVerticalAnchor constant that indicates how the specified list box is anchored vertically within its layout. Read/write.
- `AllowValueListEdits As Boolean  (read/write)`  
  Gets or sets whether the Edit List Items command is available when the user right-clicks a list box. Read/write Boolean.
- `ListItemsEditForm As String  (read/write)`  
  Gets or sets the name of the form that is displayed when the user chooses Edit List Items. Read/write String.
- `InheritValueList As Boolean  (read/write)`  
  Gets or sets whether a list box's value list is inherited from its field. Read/write Boolean.
- `LayoutID As Long  (read-only)`  
  Returns the unique identifier for the layout that contains the specified list box. Read-only Long.
- `ShowOnlyRowSourceValues As Boolean  (read/write)`  
  Gets or sets whether the list box can display values that aren't specified by the RowSource property. Read/write Boolean.
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

## Methods (7)

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
- `AddItem(Item As String, [Index As Variant])`  
  Adds a new item to the list of values displayed by the specified list box control.
    - `Item As String` (required): The display text for the new item.
    - `Index As Variant` (optional): The position of the item in the list. If this argument is omitted, the item is added to the end of the list.
- `RemoveItem(Index As Variant)`  
  Removes an item from the list of values displayed by the specified list box control.
    - `Index As Variant` (required): The item to be removed from the list, expressed as either an item number or the list item text.

## Events (14)

- `BeforeUpdate(Cancel As Integer)`
- `AfterUpdate()`
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
    - `Shift As Integer` (required): The state of the Shift, Ctrl, and Alt keys at the time of the event. If you need to test for the _Shift_ argument, you can use one of the following intrinsic constants as bit masks:<ul><li><p><b>acShiftMask</b> The bit mask for the Shift key.</p></li><li><p><b>acCtrlMask</b> The bit mask for the Ctrl key.</p></li><li><p><b>acAltMask</b> The bit mask for the Alt key.</p></li></ul>
- `KeyPress(KeyAscii As Integer)`  
  The KeyPress event occurs when the user presses and releases a key or key combination that corresponds to an ANSI code while a form or control has the focus. This event also occurs if you send an ANSI keystroke to a form or control by using the SendKeys action in a macro or the SendKeys statement in Visual Basic.
    - `KeyAscii As Integer` (required): Returns a numeric ANSI key code. The _KeyAscii_ argument is passed by reference; changing it sends a different character to the object. Setting the _KeyAscii_ argument to 0 cancels the keystroke so that the object doesn't recognize that a key was pressed.
- `KeyUp(KeyCode As Integer, Shift As Integer)`  
  The KeyUp event occurs when the user releases a key while a form or control has the focus. This event also occurs if you send a keystroke to a form or control by using the SendKeys action in a macro or the SendKeys statement in Visual Basic.
    - `KeyCode As Integer` (required): A key code, such as vbKeyF1 (the F1 key) or vbKeyHome (the Home key). To specify key codes, use the intrinsic constants shown in the Object Browser. You can prevent an object from receiving a keystroke by setting _KeyCode_ to 0.
    - `Shift As Integer` (required): The state of the Shift, Ctrl, and Alt keys at the time of the event. If you need to test for the _Shift_ argument, you can use one of the following intrinsic constants as bit masks:<ul><li><p><b>acShiftMask</b> The bit mask for the Shift key.</p></li><li><p><b>acCtrlMask</b> The bit mask for the Ctrl key.</p></li><li><p><b>acAltMask</b> The bit mask for the Alt key.</p></li></ul>
