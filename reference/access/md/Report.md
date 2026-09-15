# Report

**Type:** Class  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {FF240263-AF0A-432D-A544-A721E75738F8}  

A Report object refers to a particular Microsoft Access report.

**Remarks:** A Report object is a member of the Reports collection, which is a collection of all currently open reports. Within the Reports collection, individual reports are indexed beginning with zero. Refer to an individual Report object in the Reports collection either by referring to the report by name, or by referring to its index within the collection. If the report name includes a space, the name must be surrounded by brackets ([ ]).

**Example:**

```vba
Private Sub Report_NoData(Cancel As Integer)

    'Add code here that will be executed if no data
    'was returned by the Report's RecordSource
    MsgBox "No customers ordered this product this month. " & _
        "The report will now close."
    Cancel = True

End Sub
```

## Properties (143)

- `RecordSource As String  (read/write)`  
  Use the RecordSource property to specify the source of the data for a report. Read/write String.
- `Filter As String  (read/write)`  
  Use the Filter property to specify a subset of records to be displayed when a filter is applied to a form, report, query, or table. Read/write String.
- `FilterOn As Boolean  (read/write)`  
  Use the FilterOn property to specify or determine whether the Filter property for a form or report is applied. Read/write Boolean.
- `OrderBy As String  (read/write)`  
  Use the OrderBy property to specify how you want to sort records in a report. Read/write String.
- `OrderByOn As Boolean  (read/write)`  
  Use the OrderByOn property to specify whether an object's OrderBy property setting is applied. Read/write Boolean.
- `ServerFilter As String  (read/write)`  
  Use the ServerFilter property to specify a subset of records to be displayed when a server filter is applied to a report within a Microsoft Access project (.adp) or database. Read/write String.
- `Caption As String  (read/write)`  
  Gets or sets the title of the report in Print Preview. Read/write String.
- `RecordLocks As Byte  (read/write)`  
  Use the RecordLocks property to determine how records are locked and what happens when two users try to edit the same record at the same time. Read/write.
- `PageHeader As Byte  (read/write)`  
  Use the PageHeader property to specify whether a report's page header is printed on the same page as a report header. Read/write Byte.
- `PageFooter As Byte  (read/write)`  
  Use the PageFooter property to specify whether a report's page footer is printed on the same page as a report footer. Read/write Byte.
- `DateGrouping As Byte  (read/write)`  
  Use the DateGrouping property to specify how you want to group dates in a report. Read/write Byte.
- `GrpKeepTogether As Byte  (read/write)`  
  Use the GrpKeepTogether property to specify whether groups in a multiple column report that have their KeepTogether property for a group set to Whole Group or With First Detail will be kept together by page or by column. Read/write Byte.
- `Width As Integer  (read/write)`  
  Gets or sets the width of the specified object in twips. Read/write Integer.
- `Picture As String  (read/write)`  
  Use the Picture property to specify a bitmap or other type of graphic to be used as a background picture on a report. Read/write String.
- `PictureType As Byte  (read/write)`  
  Use the PictureType property to specify whether Microsoft Access stores an object's picture as a linked or an embedded object. Read/write Byte.
- `PictureSizeMode As Byte  (read/write)`  
  Use the PictureSizeMode property to specify how a picture for a form or report is sized. Read/write Byte.
- `PictureAlignment As Byte  (read/write)`  
  Use the PictureAlignment property to specify where a background picture will appear in an image control or on a form or report. Read/write Byte.Read/write.
- `PictureTiling As Boolean  (read/write)`  
  Use the PictureTiling property to specify whether a background picture is tiled across the entire image control, Form window, form, or page of a report. Read/write Boolean.
- `PicturePages As Byte  (read/write)`  
  Use the PicturePages property to specify on which page or pages of a report a picture will be displayed. Read/write Byte.
- `MenuBar As String  (read/write)`  
  Specifies a custom menu to display for a report. Read/write String.
- `Toolbar As String  (read/write)`  
  Specifies a custom toolbar to display for a report. Read/write String.
- `ShortcutMenuBar As String  (read/write)`  
  Use the ShortcutMenuBar property to specify the shortcut menu that appears when you right-click the specified object. Read/write String.
- `GridX As Integer  (read/write)`  
  Use the GridX property (along with the GridY property) to specify the horizontal and vertical divisions of the alignment grid in report Design view. Read/write Integer.
- `GridY As Integer  (read/write)`  
  Use the GridY property (along with the GridX property) to specify the horizontal and vertical divisions of the alignment grid in report Design view. Read/write Integer.
- `LayoutForPrint As Boolean  (read/write)`  
  Use the LayoutForPrint property to specify whether the report uses printer or screen fonts. Read/write Boolean.
- `FastLaserPrinting As Boolean  (read/write)`  
  Use the FastLaserPrinting property to specify whether lines and rectangles are replaced by text character lines-similar to the underscore ( _ ) and vertical bar ( | ) characters-when you print a report by using most laser printers. Replacing lines and rectangles with text character lines can make printing much faster. Read/write Boolean.
- `HelpFile As String  (read/write)`  
  The name of a help file associated with a report. Read/write String.
- `HelpContextId As Long  (read/write)`  
  The HelpContextID property specifies the context ID of a topic in the custom Help file specified by the HelpFile property setting. Read/write Long.
- `Hwnd As Long  (read/write)`  
  Use the hWnd property to determine the handle (a unique Long Integer value) assigned by Windows to the current window. Read/write Long.
- `Count As Integer  (read/write)`  
  Use the Count property to determine the number of items in a specified collection. Read-only Integer.
- `Page As Long  (read/write)`  
  The Page property specifies the current page number when a report is being printed. Read/write Long.
- `Pages As Integer  (read/write)`  
  Use the Pages property to return information needed to print page numbers in a report. Read/write Integer.
- `HasData As Long  (read/write)`  
  Use the HasData property to determine if a report is bound to an empty recordset. Read/write Long.
- `Left As Long  (read/write)`  
  Use the Left property to specify an object's location on a form or report. Read/write Long.
- `Top As Long  (read/write)`  
  Use the Top property to specify an object's location on a form or report. Read/write Long.
- `Height As Long  (read/write)`  
  Gets or sets the height of the specified object in twips. Read/write Long.
- `PrintSection As Boolean  (read/write)`  
  The PrintSection property specifies whether a section should be printed. Read/write Boolean.
- `NextRecord As Boolean  (read/write)`  
  The NextRecord property specifies whether a section should advance to the next record. Read/write Boolean.
- `MoveLayout As Boolean  (read/write)`  
  The MoveLayout property specifies whether Microsoft Access should move to the next printing location on the page. Read/write Boolean.
- `FormatCount As Integer  (read/write)`  
  Use the FormatCount property to determine the number of times the OnFormat property has been evaluated for the current section on a report. Read/write Integer.
- `PrintCount As Integer  (read/write)`  
  Use the PrintCount property to identify the number of times the OnPrint property has been evaluated for the current section of a report. Read/write Integer.
- `Visible As Boolean  (read/write)`  
  Returns or sets whether the object is visible. Read/write Boolean.
- `Painting As Boolean  (read/write)`  
  Use the Painting property to specify whether a report is repainted. Read/write Boolean.
- `PrtMip As Variant  (read/write)`  
  Use the PrtMip property in Visual Basic to set or return the device mode information specified for a form or report in the Print dialog box.
- `PrtDevMode As Variant  (read/write)`  
  Use the PrtDevMode property to set or return printing device mode information specified for a form or report in the Print dialog box. Read/write Variant.
- `PrtDevNames As Variant  (read/write)`  
  Use the PrtDevNames property to set or return information about the printer selected in the Print dialog box for a form or report. Read/write Variant.
- `ForeColor As Long  (read/write)`  
  Use the ForeColor property to specify the color for text in a control. Read/write Long.
- `CurrentX As Single  (read/write)`  
  Use the CurrentX property (along with the CurrentY property) to specify the horizontal and vertical coordinates for the starting position of the next printing and drawing method on a report. Read/write Single.
- `CurrentY As Single  (read/write)`  
  Use the CurrentY property (along with the CurrentX property) to specify the horizontal and vertical coordinates for the starting position of the next printing and drawing method on a report. Read/write Single.
- `ScaleHeight As Single  (read/write)`  
  Use the ScaleHeight property to specify the number of units for the vertical measurement of the page when the Circle, Line, Pset, or Print method is used while a report is printed or previewed, or its output is saved to a file. Read/write Single.
- `ScaleLeft As Single  (read/write)`  
  Use the ScaleLeft property to specify the units for the horizontal coordinates that describe the location of the left edge of a page when the Circle, Line, Pset, or Print method is used while a report is printed or previewed, or its output is saved to a file. Read/write Single.
- `ScaleMode As Integer  (read/write)`  
  Use the ScaleMode property in Visual Basic to specify the unit of measurement for coordinates on a page when the Circle, Line, Pset, or Print method is used while a report is previewed or printed, or its output is saved to a file. Read/write Integer.
- `ScaleTop As Single  (read/write)`  
  Use the ScaleTop property to specify the units for the vertical coordinates that describe the location of the top edge of a page when the Circle, Line, Pset, or Print method is used while a report is printed or previewed, or its output is saved to a file. Read/write Single.
- `ScaleWidth As Single  (read/write)`  
  Use the ScaleWidth property to specify the number of units for the horizontal measurement of the page when the Circle, Line, Pset, or Print method is used while a report is printed or previewed, or its output is saved to a file. Read/write Single.
- `FontBold As Integer  (read/write)`  
  Use the FontBold property to specify whether a font appears in a bold style in the following situations:
- `FontItalic As Integer  (read/write)`  
  Use the FontItalic property to specify whether text is italic in the following situations:
- `FontName As String  (read/write)`  
  Use the FontName property to specify the font for text in the following situations:
- `FontSize As Integer  (read/write)`  
  Use the FontSize property to specify the point size for text in the following situations:
- `FontUnderline As Integer  (read/write)`  
  Use the FontUnderline property to specify whether text is underlined in the following situations:
- `DrawMode As Integer  (read/write)`  
  Use the DrawMode property to specify how the pen (the color used in drawing) interacts with existing background colors on a report when the Line, Circle, or Pset method is used to draw on a report when printing. Read/write Integer.
- `DrawStyle As Integer  (read/write)`  
  Use the DrawStyle property to specify the line style when using the Line and Circle methods to print lines on reports. Read/write Integer.
- `DrawWidth As Integer  (read/write)`  
  Use the DrawWidth property to specify the line width for the Line, Circle, and Pset methods to print lines on reports. Read/write Integer.
- `FillColor As Long  (read/write)`  
  You use the FillColor property to specify the color that fills in boxes and circles drawn on reports with the Line and Circle methods. You can also use this property with Visual Basic to create special visual effects on custom reports when you print by using a color printer or preview the reports on a color monitor. Read/write Long.
- `FillStyle As Integer  (read/write)`  
  Use the FillStyle property to specify whether a circle or line drawn by the Circle or Line method on a report is transparent, opaque, or filled with a pattern. Read/write Integer.
- `PaletteSource As String  (read/write)`  
  Use the PaletteSource property to specify the palette for a report. Read/write String.
- `Tag As String  (read/write)`  
  Stores extra information about a form, report, section, or control needed by a Microsoft Access application. Read/write String.
- `PaintPalette As Variant  (read/write)`  
  Use the PaintPalette property to specify a palette to be used by a report. Read/write Variant.
- `OnOpen As String  (read/write)`  
  Sets or returns the value of the On Open box in the Properties window of a form or report. Read/write String.
- `OnClose As String  (read/write)`  
  Sets or returns the value of the On Close box in the Properties window of a form or report. Read/write String.
- `OnActivate As String  (read/write)`  
  Sets or returns the value of the On Activate box in the Properties window of a form or report. Read/write String.
- `OnDeactivate As String  (read/write)`  
  Sets or returns the value of the On Deactivate box in the Properties window of a form or report. Read/write String.
- `OnNoData As String  (read/write)`  
  Sets or returns the value of the On No Data box in the Properties window of a report. Read/write String.
- `OnPage As String  (read/write)`  
  Sets or returns the value of the On Page box in the Properties window of a report. Read/write String.
- `OnError As String  (read/write)`  
  Sets or returns the value of the On Error box in the Properties window of a form or report. Read/write String.
- `Dirty As Boolean  (read/write)`  
  Use the Dirty property to determine whether the current record has been modified since it was last saved. Read/write Boolean.
- `CurrentRecord As Long  (read/write)`  
  Use the CurrentRecord property to identify the current record in the recordset being viewed. Read/write Long.
- `PictureData As Variant  (read/write)`  
  Use the PictureData property to copy the picture to another object that supports the Picture property. Read/write Variant.
- `PicturePalette As Variant  (read/write)`  
  Use the PicturePalette property to specify a palette to be used by a report. Read/write Variant.
- `HasModule As Boolean  (read/write)`  
  Use the HasModule property to specify or determine whether a form or report has a class module. Read/write Boolean.
- `Orientation As Byte  (read/write)`  
  Use the Orientation property to specify or determine the view orientation. Read/write Byte.
- `InputParameters As String  (read/write)`  
  Use the InputParameters property to specify or determine the input parameters that are passed to a SQL statement in the RecordSource property of a form or report or a stored procedure when used as the record source within a Microsoft Access project (.adp). Read/write String.
- `Application As Application  (read-only)`  
  Use the Application property to access the active Microsoft Access Application object and its related properties. Read-only Application object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `ActiveControl As Control  (read-only)`  
  Use the ActiveControl property together with the Screen object to identify or refer to the control that has the focus. Read-only Control object.
- `DefaultControl As Control  (read-only)`  
  The DefaultControl property returns a Control object with which you can set the default properties for a particular type of control on a particular report. Read-only.
- `GroupLevel As GroupLevel  (read-only)`  
  Use the GroupLevel property in Visual Basic to refer to the group level that you are grouping or sorting on in a report. Read-only GroupLevel object.
- `Report As Report  (read-only)`  
  Use the Report property to refer to a report or to refer to the report associated with a subreport control. Read-only Report.
- `Module As Module  (read-only)`  
  Use the Module property to specify a report module. Read-only Module object.
- `Properties As Properties  (read-only)`  
  Returns a reference to a control's Properties collection object. Read-only.
- `Controls As Controls  (read-only)`  
  Returns the Controls collection of a form, subform, report, or section. Read-only Controls.
- `Name As String  (read/write)`  
  Use the Name property to specify or determine the string expression that identifies the name of an object. Read/write String.
- `AutoResize As Boolean  (read/write)`  
  Returns or sets a Boolean indicating whether a Report window opens automatically sized to display complete records. Read/write.
- `AutoCenter As Boolean  (read/write)`  
  Returns or sets a Boolean indicating whether a report will be centered automatically in the application window when the form is opened. Read/write.
- `PopUp As Boolean  (read/write)`  
  Specifies whether a report opens as a pop-up window. Read/write Boolean.
- `Modal As Boolean  (read/write)`  
  Use the Modal property to specify whether a report opens as a modal window. When a report opens as a modal window, you must close the window before you can move the focus to another object. Read/write Boolean.
- `BorderStyle As Byte  (read/write)`  
  Specifies how a control's border appears. Read/write Byte.
- `ControlBox As Boolean  (read/write)`  
  Specifies whether a report has a Control menu in Report view. Read/write Boolean.
- `MinMaxButtons As Byte  (read/write)`  
  Use the MinMaxButtons property to specify whether the Maximize and Minimize buttons will be visible on a report. Read/write Byte.
- `CloseButton As Boolean  (read/write)`  
  Specifies whether the Close button on a form is enabled. Read/write Boolean.
- `WindowWidth As Integer  (read/write)`  
  Returns the width of a report in twips. Read-only Integer.
- `WindowHeight As Integer  (read/write)`  
  Returns the height of a report in twips. Read-only Integer.
- `WindowTop As Integer  (read-only)`  
  Returns an Integer indicating the screen position in twips of the top edge of a report relative to the top of the Microsoft Access window. Read-only.
- `WindowLeft As Integer  (read-only)`  
  Returns an Integer indicating the screen position in twips of the left edge of a report relative to the left edge of the Microsoft Access window. Read-only.
- `OpenArgs As Variant  (read/write)`  
  Determines the string expression specified by the OpenArgs argument of the OpenReport method that opened a report. Read/write Variant.
- `Printer As _Printer  (read/write)`  
  Returns or sets a Printer object representing the default printer on the current system. Read/write.
- `Moveable As Boolean  (read/write)`  
  Returns or sets a Boolean indicating whether the specified report can be moved by the user; True if it can be moved. Read/write.
- `UseDefaultPrinter As Boolean  (read/write)`  
  Returns or sets a Boolean indicating whether the specified report uses the default printer for the system; True if the form or report uses the default printer. Read/write.
- `Recordset As Object  (read/write)`  
  Returns or sets the ADO Recordset or DAO Recordset object that represents the record source for the specified object. Read/write Object.
- `RecordSourceQualifier As String  (read/write)`  
  Returns or sets a String indicating the SQL Server owner name of the record source for the specified report. Read/write.
- `Shape As String  (read-only)`  
  Returns a String representing the shape command corresponding to the sorting and grouping of the specified report. Read-only.
- `FilterOnLoad As Boolean  (read/write)`  
  Gets or sets whether the filter specified by the Filter property is applied when the report is loaded. Read/write Boolean.
- `OrderByOnLoad As Boolean  (read/write)`  
  Gets or sets whether the sorting specified by the OrderBy property is applied when the report is loaded. Read/write Boolean.
- `DefaultView As Byte  (read/write)`  
  Use the DefaultView property to specify the opening view of a report. Read/write Byte.
- `AllowReportView As Boolean  (read/write)`  
  Gets or sets whether the user is allowed to enter Report view while using the specified report. Read/write Boolean.
- `ScrollBars As Byte  (read/write)`  
  Gets or sets whether scroll bars appear on a report. Read/write Byte.
- `Cycle As Byte  (read/write)`  
  Use the Cycle property to specify what happens when you press the Tab key and the focus is in the last control on a report. Read/write Byte.
- `OnCurrent As String  (read/write)`  
  Sets or returns the value of the OnCurrent property on the report. Read/write String.
- `KeyPreview As Boolean  (read/write)`  
  Use the KeyPreview property to specify whether the report-level keyboard event procedures are invoked before a control's keyboard event procedures. Read/write Boolean.
- `TimerInterval As Long  (read/write)`  
  Use the TimerInterval property to specify the interval, in milliseconds, between Timer events on a report. Read/write Long.
- `CurrentView As Integer  (read/write)`  
  Use the CurrentView property to determine how a report is currently displayed. Read/write Integer.
- `ShowPageMargins As Boolean  (read/write)`  
  Gets or sets whether page margins are displayed when the specified report is in Layout view. Read/write Boolean.
- `FitToPage As Boolean  (read/write)`  
  Gets or sets whether the width of the specified report is sized to automatically fit the page. Read/write Boolean.
- `AllowLayoutView As Boolean  (read/write)`  
  Gets or sets whether the specified report can be used in Layout view. Read/write Boolean.
- `OnLoad As String  (read/write)`  
  Sets or returns the value of the On Load box in the Properties window of a report. Read/write String.
- `OnResize As String  (read/write)`  
  Sets or returns the value of the On Resize box in the Properties window of a report. Read/write String.
- `OnUnload As String  (read/write)`  
  Sets or returns the value of the On Unload box in the Properties window of a form. Read/write String.
- `OnGotFocus As String  (read/write)`  
  Sets or returns the value of the On Got Focus box in the Properties window of the specified report. Read/write String.
- `OnLostFocus As String  (read/write)`  
  Sets or returns the value of the On Lost Focus box in the Properties window of the specified report. Read/write String.
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
- `OnFilter As String  (read/write)`  
  Sets or returns the value of the On Filter box in the Properties window of a report. Read/write String.
- `OnApplyFilter As String  (read/write)`  
  Sets or returns the value of the On Apply Filter box in the Properties window of a report. Read/write String.
- `OnTimer As String  (read/write)`  
  Sets or returns the value of the On Timer box in the Properties window of a form. Read/write String.
- `MouseWheel As String  (read/write)`  
  Returns or sets a String indicating which macro, event procedure, or user-defined function runs when the MouseWheel event occurs. Read/write.
- `DisplayOnSharePointSite As Byte  (read/write)`  
  Gets or sets whether the specified report can be made available as a view on a Microsoft SharePoint Foundation site. Read/write Byte.
- `Section As _Section  (read-only)`  
  Use the Section property to identify a section of a report and provide access to the properties of that section. Read-only Section object.
- `RibbonName As String  (read/write)`  
  Gets or sets the name of the customized ribbon to be displayed when the specified report is loaded. Read/write String.

## Methods (9)

- `Circle(flags As Integer, X As Single, Y As Single, radius As Single, color As Long, start As Single, end As Single, aspect As Single)`  
  The Circle method draws a circle, an ellipse, or an arc on a Report object when the Print event occurs.
    - `radius As Single` (required): Indicates the radius of the circle, ellipse, or arc. The Scale properties (ScaleMode, ScaleLeft, ScaleTop, ScaleHeight, and ScaleWidth) of the Report object specified by the Object argument determine the unit of measure used. By default, distances are measured in twips.
    - `color As Long` (required): Indicates the RGB (red-green-blue) color of the circle outline. If this argument is omitted, the value of the ForeColor property is used. You can also use the RGB function or QBColor function to specify the color.
    - `start As Single` (required): When a partial circle or ellipse is drawn, the Start argument specifies (in radians) the beginning position of the arc. The default value for the Start argument is 0 radians. The range is -2 pi radians to 2 pi radians.
    - `end As Single` (required): When a partial circle or ellipse is drawn, the End argument specifies (in radians) the end position of the arc. The default value for the End argument is 2 pi radians. The range is -2 pi radians to 2 pi radians.
    - `aspect As Single` (required): Indicates the aspect ratio of the circle. The default value is 1.0, which yields a perfect (nonelliptical) circle on any screen.
- `Line(flags As Integer, x1 As Single, y1 As Single, x2 As Single, y2 As Single, color As Long)`  
  The Line method draws lines and rectangles on a Report object when the Print event occurs.
    - `color As Long` (required): Indicates the RGB (red-green-blue) color used to draw the line. If this argument is omitted, the value of the ForeColor property is used. You can also use the RGB function or QBColor function to specify the color.
- `PSet(flags As Integer, X As Single, Y As Single, color As Long)`  
  The PSet method sets a point on a Report object to a specified color when the Print event occurs.
    - `flags As Integer` (required): A keyword that indicates that the coordinates are relative to the current graphics position given by the settings for the CurrentX and CurrentY properties of the Object argument.
    - `X As Single` (required): The horizontal coordinate of the point to set.
    - `Y As Single` (required): The vertical coordinate of the point to set.
    - `color As Long` (required): The RGB (red-green-blue) color to set the point to. If this argument is omitted, the value of the ForeColor property is used. You can also use the RGB function or QBColor function to specify the color.
- `Scale(flags As Integer, x1 As Single, y1 As Single, x2 As Single, y2 As Single)`  
  The Scale method defines the coordinate system for a Report object.
    - `x1 As Single` (required): A value for the horizontal coordinate that defines the position of the upper-left corner of the object.
    - `y1 As Single` (required): A value for the vertical coordinate that defines the position of the upper-left corner of the object.
    - `x2 As Single` (required): A value for the horizontal coordinate that defines the position of the lower-right corner of the object.
    - `y2 As Single` (required): A value for the vertical coordinate that defines the position of the lower-right corner of the object.
- `TextWidth(Expr As String) As Single`  
  The TextWidth method returns the width of a text string as it would be printed in the current font of a Report object.
    - `Expr As String` (required): The text string for which the text width will be determined.
- `TextHeight(Expr As String) As Single`  
  The TextHeight method returns the height of a text string as it would be printed in the current font of a Report object.
    - `Expr As String` (required): The text string for which the text height will be determined.
- `Print(Expr As String)`  
  The Print method prints text on a Report object by using the current color and font.
    - `Expr As String` (required): The string expressions to print. If this argument is omitted, the Print method prints a blank line. Multiple expressions can be separated with a space, a semicolon (;), or a comma. A space has the same effect as a semicolon.
- `Move(Left As Variant, [Top As Variant], [Width As Variant], [Height As Variant])`  
  Moves the specified object to the coordinates specified by the argument values.
    - `Left As Variant` (required): The screen position in twips for the left edge of the object relative to the left edge of the Microsoft Access window.
    - `Top As Variant` (optional): The screen position in twips for the top edge of the object relative to the top edge of the Access window.
    - `Width As Variant` (optional): The desired width of the object in twips.
    - `Height As Variant` (optional): The desired height of the object in twips.
- `Requery()`  
  The Requery method updates the data underlying the specified report by requerying the source of data for the control.

## Events (25)

- `Open(Cancel As Integer)`  
  The Open event occurs before a report is previewed or printed.
    - `Cancel As Integer` (required): The setting determines if the opening of the form or report occurs. Setting the Cancel argument to True (1) cancels the opening of the form or report.
- `Close()`  
  The Close event occurs when a report is closed and removed from the screen.
- `Activate()`  
  The Activate event occurs when a report receives the focus and becomes the active window.
- `Deactivate()`  
  The Deactivate event occurs when a report loses the focus to a Table, Query, Form, Report, Macro, or Module window, or to the Database window.
- `Error(DataErr As Integer, Response As Integer)`  
  The Error event occurs when a run-time error is produced in Microsoft Access when a report has the focus.
    - `DataErr As Integer` (required): The error code returned by the Err object when an error occurs. Use the DataErr argument with the Error function to map the number to the corresponding error message.
    - `Response As Integer` (required): The setting determines whether or not an error message is displayed. The Response argument can be one of the following intrinsic constants: <ul><li><b>acDataErrContinue</b> Ignore the error and continue without displaying the default Microsoft Access error message. You can supply a custom error message in place of the default error message.</li><li><b>acDataErrDisplay</b> (Default) Display the default Access error message.</li></ul>
- `NoData(Cancel As Integer)`  
  The NoData event occurs after Microsoft Access formats a report for printing that has no data (the report is bound to an empty recordset), but before the report is printed. Use this event to cancel printing of a blank report.
    - `Cancel As Integer` (required): The setting determines whether to print the report. Setting the Cancel argument to True (1) prevents the report from printing. You can also use the CancelEvent method of the DoCmd object to cancel printing the report.
- `Page()`  
  The Page event occurs after Microsoft Access formats a page of a report for printing, but before the page is printed. Use this event to draw a border around the page, or add other graphic elements to the page.
- `Current()`  
  Occurs when the focus moves to a record, making it the current record, or when the report is refreshed or requeried.
- `Load()`  
  Occurs when a report is opened and its records are displayed.
- `Resize()`  
  The Resize event occurs when a report is opened and whenever the size of a report changes.
- `Unload(Cancel As Integer)`  
  The Unload event occurs after a report is closed but before it's removed from the screen.
    - `Cancel As Integer` (required): Set to True to cancel the Unload event.
- `GotFocus()`  
  The GotFocus event occurs when the report receives the focus.
- `LostFocus()`  
  The LostFocus event occurs when the specified report loses the focus.
- `Click()`  
  The Click event occurs when the user presses and then releases a mouse button over a report.
- `DblClick(Cancel As Integer)`  
  The DblClick event occurs when the user presses and releases the left mouse button twice over a report within the double-click time limit of the system.
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
  The KeyDown event occurs when the user presses a key while a report has the focus. This event also occurs if you send a keystroke to a report by using the SendKeys action in a macro or the SendKeys statement in Visual Basic.
    - `KeyCode As Integer` (required): A key code, such as vbKeyF1 (the F1 key) or vbKeyHome (the Home key). To specify key codes, use the intrinsic constants shown in the Object Browser. You can prevent an object from receiving a keystroke by setting KeyCode to 0.
    - `Shift As Integer` (required): The state of the Shift, Ctrl, and Alt keys at the time of the event. If you need to test for the Shift argument, you can use one of the following intrinsic constants as bit masks:<ul><li><p><b>acShiftMask</b> The bit mask for the Shift key.</p></li><li><p><b>acCtrlMask</b> The bit mask for the Ctrl key.</p></li><li><p><b>acAltMask</b> The bit mask for the Alt key.</p></li></ul>
- `KeyPress(KeyAscii As Integer)`  
  The KeyPress event occurs when the user presses and releases a key or key combination that corresponds to an ANSI code while a report has the focus. This event also occurs if you send an ANSI keystroke to a report by using the SendKeys action in a macro or the SendKeys statement in Visual Basic.
    - `KeyAscii As Integer` (required): Returns a numeric ANSI key code. The KeyAscii argument is passed by reference; changing it sends a different character to the object. Setting the KeyAscii argument to 0 cancels the keystroke so that the object doesn't recognize that a key was pressed.
- `KeyUp(KeyCode As Integer, Shift As Integer)`  
  The KeyUp event occurs when the user releases a key while a report has the focus. This event also occurs if you send a keystroke to a report by using the SendKeys action in a macro or the SendKeys statement in Visual Basic.
    - `KeyCode As Integer` (required): A key code, such as vbKeyF1 (the F1 key) or vbKeyHome (the Home key). To specify key codes, use the intrinsic constants shown in the Object Browser. You can prevent an object from receiving a keystroke by setting KeyCode to 0.
    - `Shift As Integer` (required): The state of the Shift, Ctrl, and Alt keys at the time of the event. If you need to test for the Shift argument, you can use one of the following intrinsic constants as bit masks:<ul><li><p><b>acShiftMask</b> The bit mask for the Shift key.</p></li><li><p><b>acCtrlMask</b> The bit mask for the Ctrl key.</p></li><li><p><b>acAltMask</b> The bit mask for the Alt key.</p></li></ul>
- `Timer()`  
  The Timer event occurs for a report at regular intervals as specified by the report's TimerInterval property.
- `Filter(Cancel As Integer, FilterType As Integer)`  
  Occurs when the user opens a filter window by choosing Advanced Filter/Sort.
    - `Cancel As Integer` (required): The setting determines whether to open the filter window. Setting the Cancel argument to True (1) prevents the filter window from opening. You can also use the CancelEvent method of the DoCmd object to cancel opening the filter window.
    - `FilterType As Integer` (required): The filter window the user is trying to open. The FilterType argument can be acFilterAdvanced.
- `ApplyFilter(Cancel As Integer, ApplyType As Integer)`  
  Occurs when a filter is applied to a report.
    - `Cancel As Integer` (required): The setting determines if the ApplyFilter event occurs. Setting the Cancel argument to True cancels the ApplyFilter event and the filter is not applied to the report.
    - `ApplyType As Integer` (required): Returns the type of filter that was applied.
- `MouseWheel(Page As Boolean, Count As Long)`  
  Occurs when the user rolls the mouse wheel in Report view or Layout view.
    - `Page As Boolean` (required): True if the page was changed.
    - `Count As Long` (required): The number of lines by which the view was scrolled with the mouse wheel.
