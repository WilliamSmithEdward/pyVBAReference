# Form

**Type:** Class  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {7398AAFD-6527-48C7-95B7-BEABACD1CA3F}  

A Form object refers to a particular Microsoft Access form.

**Remarks:** A Form object is a member of the Forms collection, which is a collection of all currently open forms. Within the Forms collection, individual forms are indexed beginning with zero. Refer to an individual Form object in the Forms collection either by referring to the form by name, or by referring to its index within the collection. If you want to refer to a specific form in the Forms collection, it's better to refer to the form by name because a form's collection index may change. If the form name includes a space, the name must be surrounded by brackets ([ ]). Each Form object has a Controls collection, which contains all controls on the form. Refer to a control on a form either by implicitly or explicitly referring to the Controls collection. Your code will be faster if you refer to the Controls collection implicitly. The following examples show two of the ways you might refer to a control named NewData on the form called OrderForm. The next two examples show how you might refer to a control named NewData on a subform ctlSubForm contained in the form called OrderForm.

**Example:**

```vba
Private Sub cmdSearch_Click()

   Dim db As DAO.Database
   Dim qd As QueryDef
   Dim vWhere As Variant

   Set db = CurrentDb()

   On Error Resume Next
   db.QueryDefs.Delete "Query1"
   On Error GoTo 0

   vWhere = Null

   vWhere = vWhere & " AND [PayeeID]=" + Me.cboPayeeID

   If Nz(Me.txtEndDate, "") <> "" And Nz(Me.txtStartDate, "") <> "" Then
      vWhere = vWhere & " AND [RefundProcessed] Between #" & _
      Me.txtStartDate & "# AND #" & Me.txtEndDate & "#"
   Else
      If Nz(Me.txtEndDate, "") = "" And Nz(Me.txtStartDate, "") <> "" Then
         vWhere = vWhere & " AND [RefundProcessed]>=#" _
                  + Me.txtStartDate & "#"
      Else
         If Nz(Me.txtEndDate, "") <> "" And Nz(Me.txtStartDate, "") = "" Then
            vWhere = vWhere & " AND [RefundProcessed] <=#" _
                     + Me.txtEndDate & "#"
      End If
     End If
   End If

   If Nz(vWhere, "") = "" Then
      MsgBox "There are no search criteria selected." & vbCrLf & vbCrLf & _
             "Search Cancelled.", vbInformation, "Search Canceled."
   Else
      Set qd = db.CreateQueryDef("Query1", "SELECT * FROM tblRefundData? & _
               " WHERE " & Mid(vWhere, 6))
      db.Close
      Set db = Nothing

      DoCmd.OpenQuery "Query1", acViewNormal, acReadOnly
   End If
End Sub
```

## Properties (193)

- `RecordSource As String  (read/write)`  
  Use the RecordSource property to specify the source of the data for a form. Read/write String.
- `Filter As String  (read/write)`  
  Use the Filter property to specify a subset of records to be displayed when a filter is applied to a form, report, query, or table. Read/write String.
- `FilterOn As Boolean  (read/write)`  
  Use the FilterOn property to specify or determine whether the Filter property for a form or report is applied. Read/write Boolean.
- `OrderBy As String  (read/write)`  
  Use the OrderBy property to specify how you want to sort records in a form. Read/write String.
- `OrderByOn As Boolean  (read/write)`  
  Use the OrderByOn property to specify whether an object's OrderBy property setting is applied. Read/write Boolean.
- `AllowFilters As Boolean  (read/write)`  
  Use the AllowFilters property to specify whether records in a form can be filtered. Read/write Boolean.
- `Caption As String  (read/write)`  
  Gets or sets the text that appears in the title bar in Form view. Read/write String.
- `DefaultView As Byte  (read/write)`  
  Use the DefaultView property to specify the opening view of a form. Read/write Byte.
- `ViewsAllowed As Byte  (read/write)`  
  Use the ViewsAllowed property to specify whether users can switch between Datasheet view and Form view by choosing the Form view or Datasheet view command on the View menu, or by choosing the arrow next to the View button and choosing Form view or Datasheet view. Read/write Byte.
- `AllowEdits As Boolean  (read/write)`  
  Use the AllowEdits property to specify whether a user can edit saved records when using a form. Read/write Boolean.
- `AllowDeletions As Boolean  (read/write)`  
  Use the AllowDeletions property to specify whether a user can delete a record when using a form. Read/write Boolean.
- `AllowAdditions As Boolean  (read/write)`  
  Use the AllowAdditions property to specify whether a user can add a record when using a form. Read/write Boolean.
- `DataEntry As Boolean  (read/write)`  
  Use the DataEntry property to specify whether a bound form opens to allow data entry only. The Data Entry property doesn't determine whether records can be added; it only determines whether existing records are displayed. Read/write Boolean.
- `RecordsetType As Byte  (read/write)`  
  Use the RecordsetType property to specify what kind of recordset is made available to a form. Read/write Byte.
- `RecordLocks As Byte  (read/write)`  
  Use the RecordLocks property to determine how records are locked and what happens when two users try to edit the same record at the same time. Read/write.
- `ScrollBars As Byte  (read/write)`  
  Use the ScrollBars property to specify whether scroll bars appear on a form. Read/write Byte.
- `RecordSelectors As Boolean  (read/write)`  
  Use the RecordSelectors property to specify whether a form displays record selectors in Form view. Read/write Boolean.
- `NavigationButtons As Boolean  (read/write)`  
  Use the NavigationButtons property to specify whether navigation buttons and a record number box are displayed on a form. Read/write Boolean.
- `DividingLines As Boolean  (read/write)`  
  Use the DividingLines property to specify whether dividing lines will separate sections on a form or records displayed on a continuous form. Read/write Boolean.
- `AutoResize As Boolean  (read/write)`  
  Returns or sets a Boolean indicating whether a Form window opens automatically sized to display complete records. Read/write.
- `AutoCenter As Boolean  (read/write)`  
  Returns or sets a Boolean indicating whether a form will be centered automatically in the application window when the form is opened. Read/write.
- `PopUp As Boolean  (read/write)`  
  Specifies whether a form opens as a pop-up window. Read/write Boolean.
- `Modal As Boolean  (read/write)`  
  Use the Modal property to specify whether a form opens as a modal window. When a form opens as a modal window, you must close the window before you can move the focus to another object. Read/write Boolean.
- `BorderStyle As Byte  (read/write)`  
  Specifies the type of border and border elements (title bar, Control menu, Minimize and Maximize buttons, or Close button) to use for the form. You typically use different border styles for normal forms, pop-up forms, and custom dialog boxes. Read/write Byte.
- `ControlBox As Boolean  (read/write)`  
  Specifies whether a form has a Control menu in Form view and Datasheet view. Read/write Boolean.
- `MinMaxButtons As Byte  (read/write)`  
  Use the MinMaxButtons property to specify whether the Maximize and Minimize buttons will be visible on a form. Read/write Byte.
- `CloseButton As Boolean  (read/write)`  
  Specifies whether the Close button on a form is enabled. Read/write Boolean.
- `Width As Integer  (read/write)`  
  Gets or sets the width of the specified object in twips. Read/write Integer.
- `Picture As String  (read/write)`  
  Use the Picture property to specify a bitmap or other type of graphic to be used as a background picture on a form. Read/write String.
- `PictureType As Byte  (read/write)`  
  Use the PictureType property to specify whether Microsoft Access stores an object's picture as a linked or an embedded object. Read/write Byte.
- `PictureSizeMode As Byte  (read/write)`  
  Use the PictureSizeMode property to specify how a picture for a form or report is sized. Read/write Byte.
- `PictureAlignment As Byte  (read/write)`  
  Use the PictureAlignment property to specify where a background picture will appear in an image control or on a form or report. Read/write Byte.Read/write.
- `PictureTiling As Boolean  (read/write)`  
  Use the PictureTiling property to specify whether a background picture is tiled across the entire image control, Form window, form, or page of a report. Read/write Boolean.
- `Cycle As Byte  (read/write)`  
  Use the Cycle property to specify what happens when you press the Tab key and the focus is in the last control on a bound form. Read/write Byte.
- `MenuBar As String  (read/write)`  
  Specifies a custom menu to display for a form. Read/write String.
- `Toolbar As String  (read/write)`  
  Specifies a custom toolbar to display for a form. Read/write String.
- `ShortcutMenu As Boolean  (read/write)`  
  Use the ShortcutMenu property to specify whether a shortcut menu is displayed when you right-click an object on a form. For example, you might want to disable a shortcut menu to prevent the user from changing the form's underlying record source by using one of the filtering commands on the form's shortcut menu. Read/write Boolean.
- `ShortcutMenuBar As String  (read/write)`  
  Use the ShortcutMenuBar property to specify the shortcut menu that appears when you right-click the specified object. Read/write String.
- `GridX As Integer  (read/write)`  
  Use the GridX property (along with the GridY property) to specify the horizontal and vertical divisions of the alignment grid in form Design view. Read/write Integer.
- `GridY As Integer  (read/write)`  
  Use the GridY property (along with the GridX property) to specify the horizontal and vertical divisions of the alignment grid in form Design view. Read/write Integer.
- `LayoutForPrint As Boolean  (read/write)`  
  Use the LayoutForPrint property to specify whether the form uses printer or screen fonts. Read/write Boolean.
- `FastLaserPrinting As Boolean  (read/write)`  
  Use the FastLaserPrinting property to specify whether lines and rectangles are replaced by text character lines, similar to the underscore ( _ ) and vertical bar ( | ) characters, when you print a form by using most laser printers. Replacing lines and rectangles with text character lines can make printing much faster. Read/write Boolean.
- `HelpFile As String  (read/write)`  
  The name of a help file associated with a form. Read/write String.
- `HelpContextId As Long  (read/write)`  
  The HelpContextID property specifies the context ID of a topic in the custom Help file specified by the HelpFile property setting. Read/write Long.
- `RowHeight As Integer  (read/write)`  
  Use the RowHeight property to specify the height of all rows in Datasheet view. Read/write Integer.
- `DatasheetFontName As String  (read/write)`  
  Use the DatasheetFontName property to specify the font used to display and print field names and data in Datasheet view. Read/write String.
- `DatasheetFontHeight As Integer  (read/write)`  
  Use the DatasheetFontHeight property to specify the font point size used to display and print field names and data in Datasheet view. Read/write Integer.
- `DatasheetFontWeight As Integer  (read/write)`  
  Use the DatasheetFontWeight property to specify the line width of the font used to display and print characters for field names and data in Datasheet view. Read/write Integer.
- `DatasheetFontItalic As Boolean  (read/write)`  
  Use the DatasheetFontItalic property to specify an italicized appearance for field names and data in Datasheet view. Read/write Boolean.
- `DatasheetFontUnderline As Boolean  (read/write)`  
  Use the DatasheetFontUnderline property to specify an underlined appearance for field names and data in Datasheet view. Read/write Boolean.
- `DatasheetGridlinesBehavior As Byte  (read/write)`  
  Use the DatasheetGridlinesBehavior property to specify which gridlines will appear in Datasheet view. Read/write Byte.
- `DatasheetGridlinesColor As Long  (read/write)`  
  Use the DatasheetGridlinesColor property to specify the color of gridlines in a datasheet. Read/write Long.
- `DatasheetCellsEffect As Byte  (read/write)`  
  Use the DatasheetCellsEffect property to specify whether special effects are applied to cells in a datasheet. Read/write Byte.
- `DatasheetForeColor As Long  (read/write)`  
  Use the DatasheetForeColor property in Visual Basic to specify or determine the color of all text in a table, query, or form in Datasheet view within an Access database. Read/write Long.
- `DatasheetBackColor As Long  (read/write)`  
  Use the DatasheetBackColor property in Visual Basic to specify or determine the background color of an entire table, query, or form in Datasheet view within a Microsoft Access database. Read/write Long.
- `Hwnd As Long  (read/write)`  
  Use the hWnd property to determine the handle (a unique Long Integer value) assigned by Windows to the current window. Read/write Long.
- `Count As Integer  (read/write)`  
  Use the Count property to determine the number of items in a specified collection. Read-only Integer.
- `Page As Long  (read/write)`  
  The Page property specifies the current page number when a form is being printed. Read/write Long.
- `Pages As Integer  (read/write)`  
  Use the Pages property to return information needed to print page numbers in a form. Read/write Integer.
- `Visible As Boolean  (read/write)`  
  Returns or sets whether the object is visible. Read/write Boolean.
- `Painting As Boolean  (read/write)`  
  Use the Painting property to specify whether a form is repainted. Read/write Boolean.
- `PrtMip As Variant  (read/write)`  
  Use the PrtMip property in Visual Basic to set or return the device mode information specified for a form or report in the Print dialog box.
- `PrtDevMode As Variant  (read/write)`  
  Use the PrtDevMode property to set or return printing device mode information specified for a form or report in the Print dialog box. Read/write Variant.
- `PrtDevNames As Variant  (read/write)`  
  Use the PrtDevNames property to set or return information about the printer selected in the Print dialog box for a form or report. Read/write Variant.
- `FrozenColumns As Integer  (read/write)`  
  Use the FrozenColumns property to determine how many columns in a datasheet are frozen. Read/write Integer.
- `Bookmark As Variant  (read/write)`  
  Use the Bookmark property with forms to set a bookmark that uniquely identifies a particular record in the form's underlying table, query, or SQL statement. Read/write Variant.
- `PaletteSource As String  (read/write)`  
  Use the PaletteSource property to specify the palette for a form. Read/write String.
- `Tag As String  (read/write)`  
  Stores extra information about a form, report, section, or control needed by a Microsoft Access application. Read/write String.
- `PaintPalette As Variant  (read/write)`  
  Use the PaintPalette property to specify a palette to be used by a form. Read/write Variant.
- `OpenArgs As Variant  (read/write)`  
  Determines the string expression specified by the OpenArgs argument of the OpenForm method that opened a form. Read/write Variant.
- `OnCurrent As String  (read/write)`  
  Sets or returns the value of the On Current box in the Properties window of a form. Read/write String.
- `OnInsert As String  (read/write)`  
  Sets or returns the value of the Before Insert box in the Properties window of a form. Read/write String.
- `BeforeInsert As String  (read/write)`  
  Returns or sets a String indicating which macro, event procedure, or user-defined function runs when the BeforeInsert event occurs. Read/write.
- `AfterInsert As String  (read/write)`  
  Returns or sets a String indicating which macro, event procedure, or user-defined function runs when the AfterInsert event occurs. Read/write.
- `BeforeUpdate As String  (read/write)`
- `AfterUpdate As String  (read/write)`
- `OnDirty As String  (read/write)`  
  Sets or returns the value of the On Dirty box in the Properties window of a form or report. Read/write String.
- `OnDelete As String  (read/write)`  
  Sets or returns the value of the On Delete box in the Properties window of a form. Read/write String.
- `BeforeDelConfirm As String  (read/write)`  
  Returns or sets a String indicating which macro, event procedure, or user-defined function runs when the BeforeDelConfirm event occurs. Read/write.
- `AfterDelConfirm As String  (read/write)`  
  Returns or sets a String indicating which macro, event procedure, or user-defined function runs when the AfterDelConfirm event occurs. Read/write.
- `OnOpen As String  (read/write)`  
  Sets or returns the value of the On Open box in the Properties window of a form or report. Read/write String.
- `OnLoad As String  (read/write)`  
  Sets or returns the value of the On Load box in the Properties window of a form. Read/write String.
- `OnResize As String  (read/write)`  
  Sets or returns the value of the On Resize box in the Properties window of a form. Read/write String.
- `OnUnload As String  (read/write)`  
  Sets or returns the value of the On Unload box in the Properties window of a form. Read/write String.
- `OnClose As String  (read/write)`  
  Sets or returns the value of the On Close box in the Properties window of a form or report. Read/write String.
- `OnActivate As String  (read/write)`  
  Sets or returns the value of the On Activate box in the Properties window of a form or report. Read/write String.
- `OnDeactivate As String  (read/write)`  
  Sets or returns the value of the On Deactivate box in the Properties window of a form or report. Read/write String.
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
- `KeyPreview As Boolean  (read/write)`  
  Use the KeyPreview property to specify whether the form-level keyboard event procedures are invoked before a control's keyboard event procedures. Read/write Boolean.
- `OnError As String  (read/write)`  
  Sets or returns the value of the On Error box in the Properties window of a form or report. Read/write String.
- `OnFilter As String  (read/write)`  
  Sets or returns the value of the On Filter box in the Properties window of a form. Read/write String.
- `OnApplyFilter As String  (read/write)`  
  Sets or returns the value of the On Apply Filter box in the Properties window of a form. Read/write String.
- `OnTimer As String  (read/write)`  
  Sets or returns the value of the On Timer box in the Properties window of a form. Read/write String.
- `TimerInterval As Long  (read/write)`  
  Use the TimerInterval property to specify the interval, in milliseconds, between Timer events on a form. Read/write Long.
- `Dirty As Boolean  (read/write)`  
  Use the Dirty property to determine whether the current record has been modified since it was last saved. Read/write Boolean.
- `WindowWidth As Integer  (read/write)`  
  Returns the width of a form in twips. Read-only Integer.
- `WindowHeight As Integer  (read/write)`  
  Returns the height of a form in twips. Read-only Integer.
- `CurrentView As Integer  (read/write)`  
  Use the CurrentView property to determine how a form is currently displayed. Read/write Integer.
- `CurrentSectionTop As Integer  (read/write)`  
  Use this property to determine the distance in twips from the top edge of the current section to the top edge of the form. Read/write Integer.
- `CurrentSectionLeft As Integer  (read/write)`  
  Use this property to determine the distance in twips from the left side of the current section to the left side of the form. Read/write Integer.
- `SelLeft As Long  (read/write)`  
  Use the SelLeft property to specify or determine which column (field) is leftmost in the current selection rectangle. Read/write Long.
- `SelTop As Long  (read/write)`  
  Use the SelTop property to specify or determine which row (record) is topmost in the current selection rectangle in a table, query, or form datasheet, or which selected record is topmost in a continuous form. Read/write Long.
- `SelWidth As Long  (read/write)`  
  Use the SelWidth property to specify or determine the number of selected columns (fields) in the current selection rectangle. Read/write Long.
- `SelHeight As Long  (read/write)`  
  Use the SelHeight property to specify or determine the number of selected rows (records) in the current selection rectangle in a table, query, or form datasheet, or the number of selected records in a continuous form. Read/write Long.
- `CurrentRecord As Long  (read/write)`  
  Use the CurrentRecord property to identify the current record in the recordset being viewed on a form. Read/write Long.
- `PictureData As Variant  (read/write)`  
  Use the PictureData property to copy the picture to another object that supports the Picture property. Read/write Variant.
- `InsideHeight As Long  (read/write)`  
  Use the InsideHeight property (along with the InsideWidth property) to determine the height and width (in twips) of the window containing a form. Read/write Long.
- `InsideWidth As Long  (read/write)`  
  Use the InsideWidth property (along with the InsideHeight property) to determine the height and width (in twips) of the window containing a form. Read/write Long.
- `PicturePalette As Variant  (read/write)`  
  Use the PicturePalette property to specify a palette to be used by a form. Read/write Variant.
- `HasModule As Boolean  (read/write)`  
  Use the HasModule property to specify or determine whether a form or report has a class module. Read/write Boolean.
- `Orientation As Byte  (read/write)`  
  Use the Orientation property to specify or determine the view orientation. Read/write Byte.
- `ServerFilter As String  (read/write)`  
  Use the ServerFilter property to specify a subset of records to be displayed when a server filter is applied to a form within a Microsoft Access project (.adp) or database. Read/write String.
- `ServerFilterByForm As Boolean  (read/write)`  
  Use the ServerFilterByForm property to specify or determine whether a form is opened in the Server Filter By Form window. Read/write Boolean.
- `MaxRecords As Long  (read/write)`  
  Specifies the maximum number of records by a query or view. Read/write Long.
- `UniqueTable As String  (read/write)`  
  Specifies the table to be updateable when a form is bound to a multiple table view or stored procedure within a Microsoft Access project (.adp).
- `ResyncCommand As String  (read/write)`  
  Use the ResyncCommand property to specify or determine the SQL statement or stored procedure that will be used in an updateable snapshot of a table. Read/write String.
- `InputParameters As String  (read/write)`  
  Use the InputParameters property to specify or determine the input parameters that are passed to a SQL statement in the RecordSource property of a form or report, or a stored procedure when used as the record source within a Microsoft Access project (.adp). Read/write String.
- `MaxRecButton As Boolean  (read/write)`  
  Use the MaxRecButton property to specify or determine if the maximum record limit button is available on the navigation bar of a form in Datasheet view or Form view. Read/write Boolean.
- `Application As Application  (read-only)`  
  Use the Application property to access the active Microsoft Access Application object and its related properties. Read-only Application object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `NewRecord As Integer  (read-only)`  
  Use the NewRecord property to determine whether the current record is a new record. Read-only Integer.
- `ActiveControl As Control  (read-only)`  
  Use the ActiveControl property together with the Screen object to identify or refer to the control that has the focus. Read-only Control object.
- `DefaultControl As Control  (read-only)`  
  The DefaultControl property returns a Control object with which you can set the default properties for a particular type of control on a particular form. Read-only.
- `RecordsetClone As Object  (read-only)`  
  Use the RecordsetClone property to refer to a form's Recordset object specified by the form's RecordSource property. Read-only.
- `Recordset As Object  (read/write)`  
  Returns or sets the ADO Recordset or DAO Recordset object that represents the record source for the specified object. Read/write Object.
- `Form As Form  (read-only)`  
  Use the Form property to refer to a form or to refer to the form associated with a subformcontrol. Read-only Form.
- `Module As Module  (read-only)`  
  Use the Module property to specify a form module. Read-only Module object.
- `Properties As Properties  (read-only)`  
  Returns a reference to a control's Properties collection object. Read-only.
- `Controls As Controls  (read-only)`  
  Returns the Controls collection of a form, subform, report, or section. Read-only Controls.
- `Name As String  (read/write)`  
  Use the Name property to specify or determine the string expression that identifies the name of an object. Read/write String.
- `SubdatasheetHeight As Integer  (read/write)`  
  Use the SubdatasheetHeight property to specify or determine the default display height of a subdatasheet when expanded. Read/write Integer.
- `SubdatasheetExpanded As Boolean  (read/write)`  
  Use the SubdatasheetExpanded property to specify or determine the saved state of all subdatasheets within a table or query. Read/write Boolean.
- `DatasheetBorderLineStyle As Byte  (read/write)`  
  Returns or sets a Byte indicating the line style to use for the border of the specified datasheet. Read/write.
- `DatasheetColumnHeaderUnderlineStyle As Byte  (read/write)`  
  Returns or sets a Byte indicating the line style to use for the bottom edge of the column headers on the specified datasheet. Read/write.
- `HorizontalDatasheetGridlineStyle As Byte  (read/write)`  
  Returns or sets a Byte indicating the line style to use for horizontal gridlines on the specified datasheet. Read/write.
- `VerticalDatasheetGridlineStyle As Byte  (read/write)`  
  Returns or sets a Byte indicating the line style to use for vertical gridlines on the specified datasheet. Read/write.
- `WindowTop As Integer  (read-only)`  
  Returns an Integer indicating the screen position in twips of the top edge of a form relative to the top of the Microsoft Access window. Read-only.
- `WindowLeft As Integer  (read-only)`  
  Returns an Integer indicating the screen position in twips of the left edge of a form relative to the left edge of the Microsoft Access window. Read-only.
- `OnUndo As String  (read/write)`  
  Returns or sets a String indicating which macro, event procedure, or user-defined function runs when the Undo event occurs. Read/write.
- `PivotTable As Object  (read-only)`  
  Returns a PivotTable object representing a PivotTable view on a form. Read-only.
- `ChartSpace As Object  (read-only)`  
  Returns a ChartSpace object. Read-only. _Obsolete : This property uses obsolete Office Web Components and hence will no longer work.
- `Printer As _Printer  (read/write)`  
  Returns or sets a Printer object representing the default printer on the current system. Read/write.
- `Moveable As Boolean  (read/write)`  
  Returns or sets a Boolean indicating whether the specified form can be moved by the user; True if it can be moved. Read/write.
- `AllowFormView As Boolean  (read/write)`  
  Returns or sets a Boolean indicating whether the specified form may be viewed in Form view. True if Form view is allowed. Read/write.
- `AllowDatasheetView As Boolean  (read/write)`  
  Returns or sets a Boolean indicating whether the specified form may be viewed in Datasheet view. True if Datasheet view is allowed. Read/write.
- `AllowPivotTableView As Boolean  (read/write)`  
  Returns or sets a Boolean indicating whether the specified form may be viewed in PivotTable view. True if PivotTable view is allowed. Read/write.
- `AllowPivotChartView As Boolean  (read/write)`  
  Returns or sets a Boolean indicating whether the specified form may be viewed in PivotChart view. True if PivotChart view is allowed. Read/write.
- `OnConnect As String  (read/write)`  
  Returns or sets a String indicating which macro, event procedure, or user-defined function runs when the OnConnect event occurs. Read/write.
- `OnDisconnect As String  (read/write)`  
  Returns or sets a String indicating which macro, event procedure, or user-defined function runs when the OnDisconnect event occurs. Read/write.
- `PivotTableChange As String  (read/write)`  
  Returns or sets a String indicating which macro, event procedure, or user-defined function runs when the PivotTableChange event occurs. Read/write.
- `Query As String  (read/write)`  
  Returns or sets a String indicating which macro, event procedure, or user-defined function runs when the Query event occurs. Read/write.
- `BeforeQuery As String  (read/write)`  
  Returns or sets a String indicating which macro, event procedure, or user-defined function runs when the BeforeQuery event occurs. Read/write.
- `SelectionChange As String  (read/write)`  
  Returns or sets a String indicating which macro, event procedure, or user-defined function runs when the SelectionChange event occurs. Read/write.
- `CommandBeforeExecute As String  (read/write)`  
  Returns or sets a String indicating which macro, event procedure, or user-defined function runs when the CommandBeforeExecute event occurs. Read/write.
- `CommandChecked As String  (read/write)`  
  Returns or sets a String indicating which macro, event procedure, or user-defined function runs when the CommandChecked event occurs. Read/write.
- `CommandEnabled As String  (read/write)`  
  Returns or sets a String indicating which macro, event procedure, or user-defined function runs when the CommandEnabled event occurs. Read/write.
- `CommandExecute As String  (read/write)`  
  Returns or sets a String indicating which macro, event procedure, or user-defined function runs when the CommandExecute event occurs. Read/write.
- `DataSetChange As String  (read/write)`  
  Returns or sets a String indicating which macro, event procedure, or user-defined function runs when the DataSetChange event occurs. Read/write.
- `BeforeScreenTip As String  (read/write)`  
  Returns or sets a String indicating which macro, event procedure, or user-defined function runs when the BeforeScreenTip event occurs. Read/write.
- `AfterFinalRender As String  (read/write)`  
  Returns or sets a String indicating which macro, event procedure, or user-defined function runs when the AfterFinalRender event occurs. Read/write.
- `AfterRender As String  (read/write)`  
  Returns or sets a String indicating which macro, event procedure, or user-defined function runs when the AfterRender event occurs. Read/write.
- `AfterLayout As String  (read/write)`  
  Returns or sets a String indicating which macro, event procedure, or user-defined function runs when the AfterLayout event occurs. Read/write.
- `BeforeRender As String  (read/write)`  
  Returns or sets a String indicating which macro, event procedure, or user-defined function runs when the BeforeRender event occurs. Read/write.
- `MouseWheel As String  (read/write)`  
  Returns or sets a String indicating which macro, event procedure, or user-defined function runs when the MouseWheel event occurs. Read/write.
- `ViewChange As String  (read/write)`  
  Returns or sets a String indicating which macro, event procedure, or user-defined function runs when the ViewChange event occurs. Read/write.
- `DataChange As String  (read/write)`  
  Returns or sets a String indicating which macro, event procedure, or user-defined function runs when the DataChange event occurs. Read/write.
- `FetchDefaults As Boolean  (read/write)`  
  Returns or sets a Boolean indicating whether Microsoft Access shows default values for new rows on the specified form before the row is saved. True if Access shows the default values for new rows on the specified form. Read/write.
- `UseDefaultPrinter As Boolean  (read/write)`  
  Returns or sets a Boolean indicating whether the specified form uses the default printer for the system; True if the form or report uses the default printer. Read/write.
- `RecordSourceQualifier As String  (read/write)`  
  Returns or sets a String indicating the SQL Server owner name of the record source for the specified form. Read/write.
- `FilterOnLoad As Boolean  (read/write)`  
  Gets or sets whether the filter specified by the Filter property is applied when the form is loaded. Read/write Boolean.
- `OrderByOnLoad As Boolean  (read/write)`  
  Gets or sets whether the sorting specified by the OrderBy property is applied when the form is loaded. Read/write Boolean.
- `SplitFormOrientation As AcSplitFormOrientation  (read/write)`  
  Gets or sets the position of the datasheet relative to the form when the form is displayed in Split Form view. Read/write AcSplitFormOrientation.
- `SplitFormDatasheet As AcSplitFormDatasheet  (read/write)`  
  Gets or sets whether the user can edit records in the datasheet when a form is displayed in Split Form view. Read/write AcSplitFormDatasheet.
- `SplitFormSplitterBar As Boolean  (read/write)`  
  Gets or sets whether the splitter bar is available when the form is displayed in Split Form mode. Read/write Boolean.
- `SplitFormPrinting As AcSplitFormPrinting  (read/write)`  
  Gets or sets whether the contents of the form or the datasheet are printed when printing a form displayed in Split Form view. Read/write AcSplitFormPrinting.
- `SplitFormSplitterBarSave As Boolean  (read/write)`  
  Gets or sets whether the location of the splitter bar is saved when a form that's displayed in Split Form mode is closed. Read/write Boolean.
- `NavigationCaption As String  (read/write)`  
  Gets or sets the text that appears to the left of the form's navigation buttons. Read/write String.
- `AllowLayoutView As Boolean  (read/write)`  
  Gets or sets whether the specified form can be used in Layout view. Read/write Boolean.
- `DatasheetAlternateBackColor As Long  (read/write)`  
  Gets or sets the color displayed on alternate rows of a form's datasheet. Read/write Long.
- `DisplayOnSharePointSite As Byte  (read/write)`  
  Gets or sets whether the specified form can be made available as a view on a Microsoft SharePoint Foundation site. Read/write Byte.
- `SplitFormSize As Long  (read/write)`  
  Gets or sets the size in twips of the form when it's displayed in Split Form view. Read/write Long.
- `Section As _Section  (read-only)`  
  Use the Section property to identify a section of a form and provide access to the properties of that section. Read-only Section object.
- `RibbonName As String  (read/write)`  
  Gets or sets the name of the customized ribbon to be displayed when the specified form is loaded. Read/write String.
- `FitToScreen As Boolean  (read/write)`  
  Gets or sets whether the width of the form is reduced to fit the width of the screen. Read/write Boolean.

## Methods (8)

- `Undo()`  
  Use the Undo method to reset a control or form when its value has been changed.
- `Recalc()`  
  The Recalc method immediately updates all calculated controls on a form.
- `Requery()`  
  The Requery method updates the data underlying a specified form by requerying the source of data for the form.
- `Refresh()`  
  The Refresh method immediately updates the records in the underlying record source for a specified form or datasheet to reflect changes made to the data by you and other users in a multiuser environment.
- `Repaint()`  
  The Repaint method completes any pending screen updates for a specified form. When performed on a form, the Repaint method also completes any pending recalculations of the form's controls.
- `GoToPage(PageNumber As Long, [Right As Long], [Down As Long])`  
  The GoToPage method moves the focus to the first control on a specified page in the active form.
    - `PageNumber As Long` (required): A numeric expression that's a valid page number for the active form.
    - `Right As Long` (optional): A numeric expression that's a valid horizontal offset (in twips) from the left side of the window to the part of the page to be viewed.
    - `Down As Long` (optional): A numeric expression that's a valid vertical offset (in twips) from the top of the window to the part of the page to be viewed.
- `SetFocus()`  
  The SetFocus method moves the focus to the specified form, the specified control on the active form, or the specified field on the active datasheet.
- `Move(Left As Variant, [Top As Variant], [Width As Variant], [Height As Variant])`  
  Moves the specified object to the coordinates specified by the argument values.
    - `Left As Variant` (required): The screen position in twips for the left edge of the object relative to the left edge of the Microsoft Access window.
    - `Top As Variant` (optional): The screen position in twips for the top edge of the object relative to the top edge of the Access window.
    - `Width As Variant` (optional): The desired width of the object in twips.
    - `Height As Variant` (optional): The desired height of the object in twips.

## Events (50)

- `Load()`  
  Occurs when a form is opened and its records are displayed.
- `Current()`  
  Occurs when the focus moves to a record, making it the current record, or when the form is refreshed or requeried.
- `BeforeInsert(Cancel As Integer)`  
  The BeforeInsert event occurs when the user types the first character in a new record, but before the record is actually created.
    - `Cancel As Integer` (required): The setting determines if the BeforeInsert event occurs. Setting the Cancel argument to True (1) cancels the BeforeInsert event.
- `AfterInsert()`  
  The AfterInsert event occurs after a new record is added.
- `BeforeUpdate(Cancel As Integer)`
- `AfterUpdate()`
- `Delete(Cancel As Integer)`  
  Occurs when the user performs some action, such as pressing the Delete key, to delete a record, but before the record is actually deleted.
    - `Cancel As Integer` (required): The setting determines if the Delete event occurs. Setting the Cancel argument to True (1) cancels the Delete event.
- `BeforeDelConfirm(Cancel As Integer, Response As Integer)`  
  The BeforeDelConfirm event occurs after the user deletes to the buffer one or more records, but before Microsoft Access displays a dialog box asking the user to confirm the deletions.
    - `Cancel As Integer` (required): The setting determines if the BeforeDelConfirm event occurs. Setting the Cancel argument to True cancels the BeforeDelConfirm event and prevents the Delete Confirm dialog box from being displayed. If the event is canceled, the original records are restored, but the AfterDelConfirm event still occurs. If Cancel is set to True, the Response argument is ignored. If Cancel is set to False (0), which it is by default, the value in the Response argument is used by Access to determine the type of response to the Delete event.
    - `Response As Integer` (required): An intrinsic constant that determines whether Access displays the Delete Confirm dialog box asking if the record should be deleted. acDataErrContinue continues without displaying the Delete Confirm dialog box. Setting the Cancel argument to False and the Response argument to acDataErrContinue enables Microsoft Access to delete records without prompting the user. acDataErrDisplay displays the Delete Confirm dialog box. The default value is acDataErrDisplay.
- `AfterDelConfirm(Status As Integer)`  
  The AfterDelConfirm event occurs after the user confirms that the deletions and the records are actually deleted or when the deletions are canceled.
    - `Status As Integer` (required): An intrinsic constant that indicates whether a record has been deleted. acDeleteOK indicates that the deletion was successful. acDeleteCancel indicates that the deletion was canceled in Visual Basic. acDeleteUserCancel indicates that the deletion was canceled by the user.
- `Open(Cancel As Integer)`  
  The Open event occurs when a form is opened, but before the first record is displayed.
    - `Cancel As Integer` (required): The setting determines if the opening of the form or report occurs. Setting the Cancel argument to True (1) cancels the opening of the form or report.
- `Resize()`  
  The Resize event occurs when a form is opened and whenever the size of a form changes.
- `Unload(Cancel As Integer)`  
  The Unload event occurs after a form is closed but before it's removed from the screen. When the form is reloaded, Microsoft Access redisplays the form and reinitializes the contents of all its controls.
    - `Cancel As Integer` (required): Set to True to cancel the Unload event.
- `Close()`  
  The Close event occurs when a form is closed and removed from the screen.
- `Activate()`  
  The Activate event occurs when a form receives the focus and becomes the active window.
- `Deactivate()`  
  The Deactivate event occurs when a form loses the focus to a Table, Query, Form, Report, Macro, or Module window, or to the Database window.
- `GotFocus()`  
  The GotFocus event occurs when the specified object receives the focus.
- `LostFocus()`  
  The LostFocus event occurs when the specified object loses the focus.
- `Click()`  
  The Click event occurs when the user presses and then releases a mouse button over an object.
- `DblClick(Cancel As Integer)`  
  The DblClick event occurs when the user presses and releases the left mouse button twice over an object within the double-click time limit of the system.
    - `Cancel As Integer` (required): The setting determines if the DblClick event occurs. Setting the Cancel argument to True (1) cancels the DblClick event.
- `MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)`  
  The MouseDown event occurs when the user presses a mouse button.
    - `Button As Integer` (required): The button that was pressed to trigger the event. If you need to test for the Button argument, you can use one of the following intrinsic constants as bit masks:<ul><li><p><b>acLeftButton</b> The bit mask for the left mouse button.</p></li><li><p><b>acRightButton</b> The bit mask for the right mouse button.</p></li><li><p><b>acMiddleButton</b> The bit mask for the middle mouse button.</p></li></ul>
    - `Shift As Integer` (required): The state of the Shift, Ctrl, and Alt keys when the button specified by the Button argument was pressed or released. If you need to test for the Shift argument, you can use one of the following intrinsic constants as bit masks:<ul><li><p><b>acShiftMask</b> The bit mask for the Shift key.</p></li><li><p><b>acCtrlMask</b> The bit mask for the Ctrl key.</p></li><li><p><b>acAltMask</b> The bit mask for the Alt key.</p></li></ul>
    - `X As Single` (required): The x coordinate for the current location of the mouse pointer, in twips.
    - `Y As Single` (required): The y coordinate for the current location of the mouse pointer, in twips.
- `MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)`  
  The MouseMove event occurs when the user moves the mouse.
    - `Button As Integer` (required): The button that was pressed or released when the event was triggered. If you need to test for the Button argument, you can use one of the following intrinsic constants as bit masks:<ul><li><p><b>acLeftButton</b> The bit mask for the left mouse button.</p></li><li><p><b>acRightButton</b> The bit mask for the right mouse button.</p></li><li><p><b>acMiddleButton</b> The bit mask for the middle mouse button.</p></li></ul>
    - `Shift As Integer` (required): The state of the Shift, Ctrl, and Alt keys when the button specified by the Button argument was pressed or released. If you need to test for the Shift argument, you can use one of the following intrinsic constants as bit masks:<ul><li><p><b>acShiftMask</b> The bit mask for the Shift key.</p></li><li><p><b>acCtrlMask</b> The bit mask for the Ctrl key.</p></li><li><p><b>acAltMask</b> The bit mask for the Alt key.</p></li></ul>
    - `X As Single` (required): The x coordinate for the current location of the mouse pointer, in twips.
    - `Y As Single` (required): The y coordinate for the current location of the mouse pointer, in twips.
- `MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)`  
  The MouseUp event occurs when the user releases a mouse button.
    - `Button As Integer` (required): The button that was released to trigger the event. If you need to test for the Button argument, you can use one of the following intrinsic constants as bit masks:<ul><li><p><b>acLeftButton</b> The bit mask for the left mouse button.</p></li><li><p><b>acRightButton</b> The bit mask for the right mouse button.</p></li><li><p><b>acMiddleButton</b> The bit mask for the middle mouse button.</p></li></ul>
    - `Shift As Integer` (required): The state of the Shift, Ctrl, and Alt keys when the button specified by the Button argument was pressed or released. If you need to test for the Shift argument, you can use one of the following intrinsic constants as bit masks:<ul><li><p><b>acShiftMask</b> The bit mask for the Shift key.</p></li><li><p><b>acCtrlMask</b> The bit mask for the Ctrl key.</p></li><li><p><b>acAltMask</b> The bit mask for the Alt key.</p></li></ul>
    - `X As Single` (required): The x coordinate for the current location of the mouse pointer, in twips.
    - `Y As Single` (required): The y coordinate for the current location of the mouse pointer, in twips.
- `KeyDown(KeyCode As Integer, Shift As Integer)`  
  The KeyDown event occurs when the user presses a key while a form or control has the focus. This event also occurs if you send a keystroke to a form or control by using the SendKeys action in a macro or the SendKeys statement in Visual Basic.
    - `KeyCode As Integer` (required): A key code, such as vbKeyF1 (the F1 key) or vbKeyHome (the Home key). To specify key codes, use the intrinsic constants shown in the Object Browser. You can prevent an object from receiving a keystroke by setting KeyCode to 0.
    - `Shift As Integer` (required): The state of the Shift, Ctrl, and Alt keys at the time of the event. If you need to test for the Shift argument, you can use one of the following intrinsic constants as bit masks:<ul><li><p><b>acShiftMask</b> The bit mask for the Shift key.</p></li><li><p><b>acCtrlMask</b> The bit mask for the Ctrl key.</p></li><li><p><b>acAltMask</b> The bit mask for the Alt key.</p></li></ul>
- `KeyPress(KeyAscii As Integer)`  
  The KeyPress event occurs when the user presses and releases a key or key combination that corresponds to an ANSI code while a form or control has the focus. This event also occurs if you send an ANSI keystroke to a form or control by using the SendKeys action in a macro or the SendKeys statement in Visual Basic.
    - `KeyAscii As Integer` (required): Returns a numeric ANSI key code. The KeyAscii argument is passed by reference; changing it sends a different character to the object. Setting the KeyAscii argument to 0 cancels the keystroke so that the object doesn't recognize that a key was pressed.
- `KeyUp(KeyCode As Integer, Shift As Integer)`  
  The KeyUp event occurs when the user releases a key while a form or control has the focus. This event also occurs if you send a keystroke to a form or control by using the SendKeys action in a macro or the SendKeys statement in Visual Basic.
    - `KeyCode As Integer` (required): A key code, such as vbKeyF1 (the F1 key) or vbKeyHome (the Home key). To specify key codes, use the intrinsic constants shown in the Object Browser. You can prevent an object from receiving a keystroke by setting KeyCode to 0.
    - `Shift As Integer` (required): The state of the Shift, Ctrl, and Alt keys at the time of the event. If you need to test for the Shift argument, you can use one of the following intrinsic constants as bit masks:<ul><li><p><b>acShiftMask</b> The bit mask for the Shift key.</p></li><li><p><b>acCtrlMask</b> The bit mask for the Ctrl key.</p></li><li><p><b>acAltMask</b> The bit mask for the Alt key.</p></li></ul>
- `Error(DataErr As Integer, Response As Integer)`  
  The Error event occurs when a run-time error is produced in Microsoft Access when a form has the focus.
    - `DataErr As Integer` (required): The error code returned by the Err object when an error occurs. Use the DataErr argument with the Error function to map the number to the corresponding error message.
    - `Response As Integer` (required): The setting determines whether an error message is displayed. The Response argument can be one of the following intrinsic constants.<ul><li><p><b>acDataErrContinue</b> Ignore the error and continue without displaying the default Microsoft Access error message. You can supply a custom error message in place of the default error message.</p></li><li><p><b>acDataErrDisplay</b> (Default) Display the default Access error message.</p></li></ul>
- `Timer()`  
  The Timer event occurs for a form at regular intervals as specified by the form's TimerInterval property.
- `Filter(Cancel As Integer, FilterType As Integer)`  
  Occurs when the user opens a filter window by choosing Filter by Form, Advanced Filter/Sort, or Server Filter By Form.
    - `Cancel As Integer` (required): The setting determines whether to open the filter window. Setting the Cancel argument to True (1) prevents the filter window from opening. You can also use the CancelEvent method of the DoCmd object to cancel opening the filter window.
    - `FilterType As Integer` (required): The filter window the user is trying to open. The FilterType argument can be one of the following intrinsic constants:<ul><li><b>acFilterByForm</b></li><li><b>acFilterAdvanced</b></li><li><b>acServerFilterByForm</b></li></ul>
- `ApplyFilter(Cancel As Integer, ApplyType As Integer)`  
  Occurs when a filter is applied to a form.
    - `Cancel As Integer` (required): The setting determines if the ApplyFilter event occurs. Setting the Cancel argument to True cancels the ApplyFilter event and the filter is not applied to the form.
    - `ApplyType As Integer` (required): Returns the type of filter that was applied.
- `Dirty(Cancel As Integer)`  
  The Dirty event occurs when the contents of the specified control changes.
    - `Cancel As Integer` (required): The setting determines if the Dirty event occurs. Setting the Cancel argument to True (1) cancels the Dirty event.
- `Undo(Cancel As Integer)`  
  Occurs when the user undoes a change.
    - `Cancel As Integer` (required): Set this argument to True to cancel the undo operation and leave the control or form in its edited state.
- `OnConnect()`  
  Occurs when the specified PivotTable view connects to a data source.
- `OnDisconnect()`  
  Occurs when the specified PivotTable view disconnects from a data source.
- `PivotTableChange(Reason As Long)`  
  Occurs whenever the specified PivotTable view field, field set, or total is added or deleted.
    - `Reason As Long` (required): A PivotTableReasonEnum constant that indicates how the PivotTable list changed.
- `Query()`  
  Occurs whenever the specified PivotTable view query becomes necessary. The query may not occur immediately; it may be delayed until the new data is displayed.
- `BeforeQuery()`  
  Occurs when the specified PivotTable view queries its data source.
- `SelectionChange()`  
  Occurs whenever the user makes a new selection in a PivotChart view or PivotTable view.
- `CommandBeforeExecute(Command As Variant, Cancel As Object)`  
  Occurs before a specified command is executed. Use this event when you want to impose certain restrictions before a particular command is executed.
    - `Command As Variant` (required): The command that is going to be executed.
    - `Cancel As Object` (required): Set the Value property of this object to True to cancel the command.
- `CommandChecked(Command As Variant, Checked As Object)`  
  Occurs when the specified Microsoft Office web component determines whether the specified command is selected.
    - `Command As Variant` (required): The command that has been verified as being selected.
    - `Checked As Object` (required): Set the Value property of this object to False to clear the command.
- `CommandEnabled(Command As Variant, Enabled As Object)`  
  Occurs when the specified Microsoft Office web component determines whether the specified command is enabled.
    - `Command As Variant` (required): The command that has been verified as being enabled.
    - `Enabled As Object` (required): Set the Value property of this object to False to disable the command.
- `CommandExecute(Command As Variant)`  
  Occurs after the specified command is executed. Use this event when you want to execute a set of commands after a particular command is executed.
    - `Command As Variant` (required): The command that is executed.
- `DataSetChange()`  
  Occurs whenever the specified PivotTable view is data-bound and the data set changes; for example, when a filter operation takes place. This event also occurs when initial data is available from the data source.
- `BeforeScreenTip(ScreenTipText As Object, SourceObject As Object)`  
  Occurs before a ScreenTip is displayed for an element in a PivotChart view or PivotTable view.
    - `ScreenTipText As Object` (required): Set the Value property of this object to the ScreenTip that you want to display. Changing this argument to an empty string effectively hides the ScreenTip.
    - `SourceObject As Object` (required): The object that generates the ScreenTip.
- `BeforeRender(drawObject As Object, chartObject As Object, Cancel As Object)`  
  Occurs before any object in the specified PivotChart view has been rendered.
    - `drawObject As Object` (required): A reference to the ChChartDraw object. Use the DrawType property of the returned object to determine what type of rendering is about to occur.
    - `chartObject As Object` (required): The object that is to be rendered. Use the TypeName function to determine the type of the object.
    - `Cancel As Object` (required): Set the Value property of this object to True to cancel the rendering of the PivotChart view object.
- `AfterRender(drawObject As Object, chartObject As Object)`  
  Occurs after the object represented by the chartObject argument has been rendered.
    - `drawObject As Object` (required): A ChChartDraw object. Use the methods and properties of this object to draw objects on the chart.
    - `chartObject As Object` (required): The object that has just been rendered. Use the TypeName function to determine what type of object has just been rendered.
- `AfterFinalRender(drawObject As Object)`  
  Occurs after all elements in the specified PivotChart view have been rendered.
    - `drawObject As Object` (required): A ChChartDraw object. Use the methods and properties of this object to draw objects on the chart.
- `AfterLayout(drawObject As Object)`  
  Occurs after all charts in the specified PivotChart view have been laid out, but before they have been rendered.
    - `drawObject As Object` (required): A ChChartDraw object. Use the methods and properties of this object to draw objects on the chart.
- `MouseWheel(Page As Boolean, Count As Long)`  
  Occurs when the user rolls the mouse wheel in Form view, Split Form view, Datasheet view, Layout view, PivotChart view, or PivotTable view.
    - `Page As Boolean` (required): True if the page was changed.
    - `Count As Long` (required): The number of lines by which the view was scrolled with the mouse wheel.
- `ViewChange(Reason As Long)`  
  Occurs whenever the specified PivotChart view or PivotTable view is redrawn.
    - `Reason As Long` (required): A PivotViewReasonEnum constant that indicates how the view was changed. Reason always returns 1 for PivotChart views.
- `DataChange(Reason As Long)`  
  Occurs when certain properties are changed or when certain methods are executed in the specified PivotTable view.
    - `Reason As Long` (required): A PivotDataReasonEnum constant that indicates the reason that this event was triggered.
