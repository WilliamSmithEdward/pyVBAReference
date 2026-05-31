# Chart

**Type:** Class  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00020821-0000-0000-C000-000000000046}  

Represents a chart in a workbook.

**Remarks:** The chart can be either an embedded chart (contained in a ChartObject object) or a separate chart sheet. The Charts collection contains a Chart object for each chart sheet in a workbook. Use Charts (_index_), where _index_ is the chart-sheet index number or name, to return a single Chart object. The chart _index_ number represents the position of the chart sheet on the workbook tab bar. _Charts(1)_ is the first (leftmost) chart in the workbook; _Charts(Charts.Count)_ is the last (rightmost). All chart sheets are included in the index count, even if they are hidden. The chart-sheet name is shown on the workbook tab for the chart. Use the Name property of the ChartObject object to set or return the chart name. The following example changes the color of series 1 on chart sheet 1. The following example moves the chart named Sales to the end of the active workbook. The Chart object is also a member of the Sheets collection, which contains all the sheets in the workbook (both chart sheets and worksheets). Use Sheets (_index_), where _index_ is the sheet index number or name, to return a single sheet. When a chart is the active object, you can use the ActiveChart property to refer to it.

## Properties (61)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `CodeName As String  (read-only)`  
  Returns the code name for the object. Read-only String.
- `_CodeName As String  (read/write)`
- `Index As Long  (read-only)`  
  Returns a Long value that represents the index number of the object within the collection of similar objects.
- `Name As String  (read/write)`  
  Returns or sets a String value representing the name of the object.
- `Next As Object  (read-only)`  
  Returns a Worksheet object that represents the next sheet.
- `PageSetup As PageSetup  (read-only)`  
  Returns a PageSetup object that contains all the page setup settings for the specified object. Read-only.
- `Previous As Object  (read-only)`  
  Returns a Worksheet object that represents the previous sheet.
- `ProtectContents As Boolean  (read-only)`  
  True if the contents of the sheet are protected. For a chart, this protects the entire chart. To turn on content protection, use the Protect method with the _Contents_ argument set to True. Read-only Boolean.
- `ProtectDrawingObjects As Boolean  (read-only)`  
  True if shapes are protected. To turn on shape protection, use the Protect method with the _DrawingObjects_ argument set to True. Read-only Boolean.
- `ProtectionMode As Boolean  (read-only)`  
  True if user-interface-only protection is turned on. To turn on user interface protection, use the Protect method with the _UserInterfaceOnly_ argument set to True. Read-only Boolean.
- `Visible As XlSheetVisibility  (read/write)`  
  Returns or sets an XlSheetVisibility value that determines whether the object is visible.
- `Shapes As Shapes  (read-only)`  
  Returns a Shapes collection that represents all the shapes on the chart sheet. Read-only.
- `AutoScaling As Boolean  (read/write)`  
  True if Microsoft Excel scales a 3D chart so that it's closer in size to the equivalent 2D chart. The RightAngleAxes property must be True. Read/write Boolean.
- `ChartArea As ChartArea  (read-only)`  
  Returns a ChartArea object that represents the complete chart area for the chart. Read-only.
- `ChartTitle As ChartTitle  (read-only)`  
  Returns a ChartTitle object that represents the title of the specified chart. Read-only.
- `DataTable As DataTable  (read-only)`  
  Returns a DataTable object that represents the chart data table. Read-only.
- `DepthPercent As Long  (read/write)`  
  Returns or sets the depth of a 3D chart as a percentage of the chart width (between 20 and 2000 percent). Read/write Long.
- `DisplayBlanksAs As XlDisplayBlanksAs  (read/write)`  
  Returns or sets the way that blank cells are plotted on a chart. Can be one of the XlDisplayBlanksAs constants. Read/write Long.
- `Elevation As Long  (read/write)`  
  Returns or sets the elevation of the 3D chart view, in degrees. Read/write Long.
- `Floor As Floor  (read-only)`  
  Returns a Floor object that represents the floor of the 3D chart. Read-only.
- `GapDepth As Long  (read/write)`  
  Returns or sets the distance between the data series in a 3D chart as a percentage of the marker width. The value of this property must be between 0 and 500. Read/write Long.
- `HasAxis As Variant  (read/write)`  
  Returns or sets which axes exist on the chart. Read/write Variant.
- `HasDataTable As Boolean  (read/write)`  
  True if the chart has a data table. Read/write Boolean.
- `HasLegend As Boolean  (read/write)`  
  True if the chart has a legend. Read/write Boolean.
- `HasTitle As Boolean  (read/write)`  
  True if the axis or chart has a visible title. Read/write Boolean.
- `HeightPercent As Long  (read/write)`  
  Returns or sets the height of a 3D chart as a percentage of the chart width (between 5 and 500 percent). Read/write Long.
- `Hyperlinks As Hyperlinks  (read-only)`  
  Returns a Hyperlinks collection that represents the hyperlinks for the chart.
- `Legend As Legend  (read-only)`  
  Returns a Legend object that represents the legend for the chart. Read-only.
- `Perspective As Long  (read/write)`  
  Returns or sets a Long value that represents the perspective for the 3D chart view.
- `PlotArea As PlotArea  (read-only)`  
  Returns a PlotArea object that represents the plot area of a chart. Read-only.
- `PlotVisibleOnly As Boolean  (read/write)`  
  True if only visible cells are plotted. False if both visible and hidden cells are plotted. Read/write Boolean.
- `RightAngleAxes As Variant  (read/write)`  
  True if the chart axes are at right angles, independent of chart rotation or elevation. Applies only to 3D line, column, and bar charts. Read/write Boolean.
- `Rotation As Variant  (read/write)`  
  Returns or sets the rotation of the 3D chart view (the rotation of the plot area around the z-axis, in degrees). The value of this property must be from 0 to 360, except for 3D bar charts, where the value must be from 0 to 44. The default value is 20. Applies only to 3D charts. Read/write Variant.
- `ChartType As XlChartType  (read/write)`  
  Returns or sets the chart type. Read/write XlChartType.
- `Walls As Walls  (read-only)`  
  Returns a Walls object that represents the walls of the 3D chart. Read-only.
- `BarShape As XlBarShape  (read/write)`  
  Returns or sets the shape used with the 3D bar or column chart. Read/write XlBarShape.
- `PlotBy As XlRowCol  (read/write)`  
  Returns or sets the way columns or rows are used as data series on the chart. Can be one of the following XlRowCol constants: xlColumns or xlRows. Read/write Long.
- `ProtectFormatting As Boolean  (read/write)`  
  True if chart formatting cannot be modified by the user. Read/write Boolean.
- `ProtectData As Boolean  (read/write)`  
  True if series formulas cannot be modified by the user. Read/write Boolean.
- `ProtectSelection As Boolean  (read/write)`  
  True if chart elements cannot be selected. Read/write Boolean.
- `PivotLayout As PivotLayout  (read-only)`  
  Returns a PivotLayout object that represents the placement of fields in a PivotTable report and the placement of axes in a PivotChart report. Read-only.
- `Tab As Tab  (read-only)`  
  Returns a Tab object for a chart.
- `MailEnvelope As MsoEnvelope  (read-only)`  
  Represents an email header for a document.
- `ShowDataLabelsOverMaximum As Boolean  (read/write)`  
  Returns or sets whether to show the data labels when the value is greater than the maximum value on the value axis. Read/write Boolean.
- `SideWall As Walls  (read-only)`  
  Returns a Walls object that allows the user to individually format the side wall of a 3D chart. Read-only.
- `BackWall As Walls  (read-only)`  
  Returns a Walls object that allows the user to individually format the back wall of a 3D chart. Read-only.
- `ChartStyle As Variant  (read/write)`  
  Returns or sets the chart style for the chart. Read/write Variant.
- `PrintedCommentPages As Long  (read-only)`  
  Returns the number of comment pages that will be printed for the current chart. Read-only.
- `ShowReportFilterFieldButtons As Boolean  (read/write)`  
  Returns or sets whether to display the report filter field buttons on a PivotChart. Read/write.
- `ShowLegendFieldButtons As Boolean  (read/write)`  
  Returns or sets whether to display legend field buttons on a PivotChart. Read/write.
- `ShowAxisFieldButtons As Boolean  (read/write)`  
  Returns or sets whether to display axis field buttons on a PivotChart. Read/write.
- `ShowValueFieldButtons As Boolean  (read/write)`  
  Returns or sets whether to display the value field buttons on a PivotChart. Read/write.
- `ShowAllFieldButtons As Boolean  (read/write)`  
  Returns or sets whether to display all field buttons on a PivotChart. Read/write.
- `CategoryLabelLevel As XlCategoryLabelLevel  (read/write)`  
  Returns an XlCategoryLabelLevel constant referring to the level of where the category labels are being sourced from. Read/write Integer.
- `SeriesNameLevel As XlSeriesNameLevel  (read/write)`  
  Returns an XlSeriesNameLevel constant referring to the level of where the series names are being sourced from. Read/write Integer.
- `ChartColor As Variant  (read/write)`  
  Returns or sets an Integer that represents the color scheme for the chart. Read/write.
- `ShowExpandCollapseEntireFieldButtons As Boolean  (read/write)`  
  True to display the Expand Entire Field and Collapse Entire Field buttons on the specified PivotChart. Read/write Boolean.
- `DisplayValueNotAvailableAsBlank As Boolean  (read/write)`

## Methods (39)

- `Activate()`  
  Makes the current chart the active chart.
- `Copy([Before As Variant], [After As Variant])`  
  Copies the sheet to another location in the workbook.
    - `Before As Variant` (optional): The sheet before which the copied sheet will be placed. You cannot specify _Before_ if you specify _After_.
    - `After As Variant` (optional): The sheet after which the copied sheet will be placed. You cannot specify _After_ if you specify _Before_.
- `Delete()`  
  Deletes the object.
- `Move([Before As Variant], [After As Variant])`  
  Moves the chart to another location in the workbook.
    - `Before As Variant` (optional): The sheet before which the moved chart will be placed. You cannot specify _Before_ if you specify _After_.
    - `After As Variant` (optional): The sheet after which the moved chart will be placed. You cannot specify _After_ if you specify _Before_.
- `PrintPreview([EnableChanges As Variant])`  
  Shows a preview of the object as it would look when printed.
    - `EnableChanges As Variant` (optional): Pass a Boolean value to specify if the user can change the margins and other page setup options available in print preview.
- `Select([Replace As Variant])`  
  Selects the object.
    - `Replace As Variant` (optional): Used only with sheets. True to replace the current selection with the specified object. False to extend the current selection to include any previously selected objects and the specified object.
- `Unprotect([Password As Variant])`  
  Removes protection from a sheet or workbook. This method has no effect if the sheet or workbook isn't protected.
    - `Password As Variant` (optional): A string that denotes the case-sensitive password to use to unprotect the chart. If the chart isn't protected with a password, this argument is ignored.
- `Axes([Type As Variant], [AxisGroup As XlAxisGroup]) As Object`  
  Returns an object that represents either a single axis or a collection of the axes on the chart.
    - `Type As Variant` (optional): Specifies the axis to return. Can be one of the following XlAxisType constants: xlValue, xlCategory, or xlSeriesAxis (xlSeriesAxis is valid only for 3D charts).
    - `AxisGroup As XlAxisGroup` (optional): Specifies the axis group. If this argument is omitted, the primary group is used. 3D charts have only one axis group.
- `SetBackgroundPicture(Filename As String)`  
  Sets the background graphic for a chart.
    - `Filename As String` (required): The name of the graphic file.
- `ChartGroups([Index As Variant]) As Object`  
  Returns an object that represents either a single chart group (a ChartGroup object) or a collection of all the chart groups in the chart (a ChartGroups object). The returned collection includes every type of group.
    - `Index As Variant` (optional): The chart group number.
- `ChartObjects([Index As Variant]) As Object`  
  Returns an object that represents either a single embedded chart (a ChartObject object) or a collection of all the embedded charts (a ChartObjects object) on the sheet.
    - `Index As Variant` (optional): The name or number of the chart. This argument can be an array to specify more than one chart.
- `ChartWizard([Source As Variant], [Gallery As Variant], [Format As Variant], [PlotBy As Variant], [CategoryLabels As Variant], [SeriesLabels As Variant], [HasLegend As Variant], [Title As Variant], [CategoryTitle As Variant], [ValueTitle As Variant], [ExtraTitle As Variant])`  
  Modifies the properties of the given chart. Use this method to quickly format a chart without setting all the individual properties. This method is noninteractive, and it changes only the specified properties.
    - `Source As Variant` (optional): The range that contains the source data for the new chart. If this argument is omitted, Microsoft Excel edits the active chart sheet or the selected chart on the active worksheet.
    - `Gallery As Variant` (optional): One of the constants of XlChartType specifying the chart type.
    - `Format As Variant` (optional): The option number for the built-in autoformats. Can be a number from 1 through 10, depending on the gallery type. If this argument is omitted, Excel chooses a default value based on the gallery type and data source.
    - `PlotBy As Variant` (optional): Specifies whether the data for each series is in rows or columns. Can be one of the following XlRowCol constants: xlRows or xlColumns.
    - `CategoryLabels As Variant` (optional): An integer specifying the number of rows or columns within the source range that contain category labels. Legal values are from 0 (zero) through one less than the maximum number of the corresponding categories or series.
    - `SeriesLabels As Variant` (optional): An integer specifying the number of rows or columns within the source range that contain series labels. Legal values are from 0 (zero) through one less than the maximum number of the corresponding categories or series.
    - `HasLegend As Variant` (optional): True to include a legend.
    - `Title As Variant` (optional): The chart title text.
    - `CategoryTitle As Variant` (optional): The category axis title text.
    - `ValueTitle As Variant` (optional): The value axis title text.
    - `ExtraTitle As Variant` (optional): The series axis title for 3D charts or the second value axis title for 2D charts.
- `CheckSpelling([CustomDictionary As Variant], [IgnoreUppercase As Variant], [AlwaysSuggest As Variant], [SpellLang As Variant])`  
  Checks the spelling of an object.
    - `CustomDictionary As Variant` (optional): A string that indicates the file name of the custom dictionary to be examined if the word isn't found in the main dictionary. If this argument is omitted, the currently specified dictionary is used.
    - `IgnoreUppercase As Variant` (optional): True to have Microsoft Excel ignore words that are all uppercase. False to have Excel check words that are all uppercase. If this argument is omitted, the current setting will be used.
    - `AlwaysSuggest As Variant` (optional): True to have Excel display a list of suggested alternate spellings when an incorrect spelling is found. False to have Excel wait for you to input the correct spelling. If this argument is omitted, the current setting will be used.
    - `SpellLang As Variant` (optional): The language of the dictionary being used. Can be one of the MsoLanguageID values.
- `CopyPicture([Appearance As XlPictureAppearance], [Format As XlCopyPictureFormat], [Size As XlPictureAppearance])`  
  Copies the selected object to the Clipboard as a picture.
    - `Appearance As XlPictureAppearance` (optional): Specifies how the picture should be copied. The default value is xlScreen.
    - `Format As XlCopyPictureFormat` (optional): The format of the picture. The default value is xlPicture.
    - `Size As XlPictureAppearance` (optional): The size of the copied picture when the object is a chart on a chart sheet (not embedded on a worksheet). The default value is xlPrinter.
- `Evaluate(Name As Variant) As Variant`  
  Converts a Microsoft Excel name to an object or a value.
    - `Name As Variant` (required): The name of the object, using the naming convention of Microsoft Excel.
- `_Evaluate(Name As Variant) As Variant`
- `Location(Where As XlChartLocation, [Name As Variant]) As Chart`  
  Moves the chart to a new location.
    - `Where As XlChartLocation` (required): Where to move the chart.
    - `Name As Variant` (optional): Required if _Where_ is xlLocationAsObject. The name of the sheet where the chart will be embedded if _Where_ is xlLocationAsObject, or the name of the new sheet if _Where_ is xlLocationAsNewSheet.
- `OLEObjects([Index As Variant]) As Object`  
  Returns an object that represents either a single OLE object (an OLEObject) or a collection of all OLE objects (an OLEObjects collection) on the chart or sheet. Read-only.
    - `Index As Variant` (optional): The name or number of the OLE object.
- `Paste([Type As Variant])`  
  Pastes chart data from the Clipboard into the specified chart.
    - `Type As Variant` (optional): Specifies the chart information to paste if a chart is on the Clipboard. Can be one of the following XlPasteType constants: xlPasteFormats, xlPasteFormulas, or xlPasteAll. The default value is xlPasteAll. If there's data other than a chart on the Clipboard, this argument cannot be used.
- `SeriesCollection([Index As Variant]) As Object`  
  Returns an object that represents either a single series (a Series object) or a collection of all the series (a SeriesCollection collection) in the chart or chart group.
    - `Index As Variant` (optional): The name or number of the series.
- `GetChartElement(x As Long, y As Long, ElementID As Long, Arg1 As Long, Arg2 As Long)`  
  Returns information about the chart element at specified _x_ and _y_ coordinates. This method is unusual in that you specify values for only the first two arguments. Microsoft Excel fills in the other arguments, and your code should examine those values when the method returns.
    - `x As Long` (required): The _x_ coordinate of the chart element.
    - `y As Long` (required): The _y_ coordinate of the chart element.
    - `ElementID As Long` (required): When the method returns, this argument contains the XLChartItem value of the chart element at the specified coordinates. For more information, see the Remarks section.
    - `Arg1 As Long` (required): When the method returns, this argument contains information related to the chart element. For more information, see the Remarks section.
    - `Arg2 As Long` (required): When the method returns, this argument contains information related to the chart element. For more information, see the Remarks section.
- `SetSourceData(Source As Range, [PlotBy As Variant])`  
  Sets the source data range for the chart.
    - `Source As Range` (required): The range that contains the source data.
    - `PlotBy As Variant` (optional): Specifies the way the data is to be plotted. Can be either of the following XlRowCol constants: xlColumns or xlRows.
- `Export(Filename As String, [FilterName As Variant], [Interactive As Variant]) As Boolean`  
  Exports the chart in a graphic format.
    - `Filename As String` (required): The name of the exported file.
    - `FilterName As Variant` (optional): The language-independent name of the graphic filter as it appears in the registry.
    - `Interactive As Variant` (optional): True to display the dialog box that contains the filter-specific options. If this argument is False, Microsoft Excel uses the default values for the filter. The default value is False.
- `Refresh()`  
  Causes the specified chart to be redrawn immediately.
- `ApplyDataLabels([Type As XlDataLabelsType], [LegendKey As Variant], [AutoText As Variant], [HasLeaderLines As Variant], [ShowSeriesName As Variant], [ShowCategoryName As Variant], [ShowValue As Variant], [ShowPercentage As Variant], [ShowBubbleSize As Variant], [Separator As Variant])`  
  Applies data labels to all the series in a chart.
    - `Type As XlDataLabelsType` (optional): The type of data label to apply.
    - `LegendKey As Variant` (optional): True to show the legend key next to the point. The default value is False.
    - `AutoText As Variant` (optional): True if the object automatically generates appropriate text based on content.
    - `HasLeaderLines As Variant` (optional): For the Chart and Series objects, True if the series has leader lines.
    - `ShowSeriesName As Variant` (optional): Pass a Boolean value to enable or disable the series name for the data label.
    - `ShowCategoryName As Variant` (optional): Pass a Boolean value to enable or disable the category name for the data label.
    - `ShowValue As Variant` (optional): Pass a Boolean value to enable or disable the value for the data label.
    - `ShowPercentage As Variant` (optional): Pass a Boolean value to enable or disable the percentage for the data label.
    - `ShowBubbleSize As Variant` (optional): Pass a Boolean value to enable or disable the bubble size for the data label.
    - `Separator As Variant` (optional): The separator for the data label.
- `Protect([Password As Variant], [DrawingObjects As Variant], [Contents As Variant], [Scenarios As Variant], [UserInterfaceOnly As Variant])`  
  Protects a chart so that it cannot be modified.
    - `Password As Variant` (optional): A string that specifies a case-sensitive password for the worksheet or workbook. If this argument is omitted, you can unprotect the worksheet or workbook without using a password. Otherwise, you must specify the password to unprotect the worksheet or workbook. If you forget the password, you cannot unprotect the worksheet or workbook. Use strong passwords that combine uppercase and lowercase letters, numbers, and symbols. Weak passwords don't mix these elements. Strong password: Y6dh!et5. Weak password: House27. Passwords should be 8 or more characters in length. A pass phrase that uses 14 or more characters is better. It's critical that you remember your password. If you forget your password, Microsoft cannot retrieve it. Store the passwords that you write down in a secure place away from the information that they help protect.
    - `DrawingObjects As Variant` (optional): True to protect shapes. The default value is True.
    - `Contents As Variant` (optional): True to protect contents. For a chart, this protects the entire chart. For a worksheet, this protects the locked cells. The default value is True.
    - `Scenarios As Variant` (optional): True to protect scenarios. This argument is valid only for worksheets. The default value is True.
    - `UserInterfaceOnly As Variant` (optional): True to protect the user interface, but not macros. If this argument is omitted, protection applies both to macros and to the user interface.
- `ApplyLayout(Layout As Long, [ChartType As Variant])`  
  Applies the layouts shown in the ribbon.
    - `Layout As Long` (required): Specifies the type of layout. The type of layout is denoted by a number from 1 to 10.
    - `ChartType As Variant` (optional): The type of chart.
- `SetElement(Element As MsoChartElementType)`  
  Sets chart elements on a chart. Read/write MsoChartElementType.
    - `Element As MsoChartElementType` (required): Specifies the chart element type.
- `PrintOut([From As Variant], [To As Variant], [Copies As Variant], [Preview As Variant], [ActivePrinter As Variant], [PrintToFile As Variant], [Collate As Variant], [PrToFileName As Variant])`  
  Prints the object.
    - `From As Variant` (optional): The number of the page at which to start printing. If this argument is omitted, printing starts at the beginning.
    - `To As Variant` (optional): The number of the last page to print. If this argument is omitted, printing ends with the last page.
    - `Copies As Variant` (optional): The number of copies to print. If this argument is omitted, one copy is printed.
    - `Preview As Variant` (optional): True to have Microsoft Excel invoke print preview before printing the object. False (or omitted) to print the object immediately.
    - `ActivePrinter As Variant` (optional): Sets the name of the active printer.
    - `PrintToFile As Variant` (optional): True to print to a file. If _PrToFileName_ is not specified, Excel prompts the user to enter the name of the output file.
    - `Collate As Variant` (optional): True to collate multiple copies.
    - `PrToFileName As Variant` (optional): If _PrintToFile_ is set to True, this argument specifies the name of the file that you want to print to.
- `ApplyChartTemplate(Filename As String)`  
  Applies a standard or custom chart type to a chart.
    - `Filename As String` (required): The file name for a chart template.
- `SaveChartTemplate(Filename As String)`  
  Saves a custom chart template to the list of available chart templates.
    - `Filename As String` (required): The name of the chart template.
- `SetDefaultChart(Name As Variant)`  
  Specifies the name of the chart template that Microsoft Excel uses when creating new charts.
    - `Name As Variant` (required): Specifies the name of the default chart template that will be used when creating new charts. This name can be a string naming a chart in the gallery for a user-defined template, or it can be a special constant xlBuiltIn (XlChartGallery) to specify a built-in chart template.
- `ClearToMatchStyle()`  
  Clears the chart elements formatting to automatic.
- `FullSeriesCollection([Index As Variant]) As Object`  
  Enables retrieving the filtered out series specified by the _Index_ argument.
    - `Index As Variant` (optional): The indexed number of the filtered out Series object.
- `ClearToMatchColorStyle()`  
  Clears all colors on the specified chart that don't follow the color style applied to the chart.
- `ExportAsFixedFormat(Type As XlFixedFormatType, [Filename As Variant], [Quality As Variant], [IncludeDocProperties As Variant], [IgnorePrintAreas As Variant], [From As Variant], [To As Variant], [OpenAfterPublish As Variant], [FixedFormatExtClassPtr As Variant], [WorkIdentity As Variant])`  
  Exports to a file of the specified format.
    - `Type As XlFixedFormatType` (required): The type of file format to export to.
    - `Filename As Variant` (optional): The file name of the file to be saved. You can include a full path, or Microsoft Excel saves the file in the current folder.
    - `Quality As Variant` (optional): Optional XlFixedFormatQuality. Specifies the quality of the published file.
    - `IncludeDocProperties As Variant` (optional): True to include the document properties; otherwise, False.
    - `IgnorePrintAreas As Variant` (optional): True to ignore any print areas set when publishing; otherwise, False.
    - `From As Variant` (optional): The number of the page at which to start publishing. If this argument is omitted, publishing starts at the beginning.
    - `To As Variant` (optional): The number of the last page to publish. If this argument is omitted, publishing ends with the last page.
    - `OpenAfterPublish As Variant` (optional): True to display the file in the viewer after it is published; otherwise, False.
    - `FixedFormatExtClassPtr As Variant` (optional): Pointer to the FixedFormatExt class.
- `SetProperty(ID As String, Value As Variant)`
- `GetProperty(ID As String) As Variant`
- `SaveAs(Filename As String, [FileFormat As Variant], [Password As Variant], [WriteResPassword As Variant], [ReadOnlyRecommended As Variant], [CreateBackup As Variant], [AddToMru As Variant], [TextCodepage As Variant], [TextVisualLayout As Variant], [Local As Variant])`  
  Saves changes to the chart or worksheet in a different file.
    - `Filename As String` (required): Variant. A string that indicates the name of the file to be saved. You can include a full path; if you don't, Microsoft Excel saves the file in the current folder.
    - `FileFormat As Variant` (optional): The file format to use when you save the file. For a list of valid choices, see the FileFormat property. For an existing file, the default format is the last file format specified; for a new file, the default is the format of the version of Excel being used.
    - `Password As Variant` (optional): A case-sensitive string (no more than 15 characters) that indicates the protection password to be given to the file.
    - `WriteResPassword As Variant` (optional): A string that indicates the write-reservation password for this file. If a file is saved with the password and the password isn't supplied when the file is opened, the file is opened as read-only.
    - `ReadOnlyRecommended As Variant` (optional): True to display a message when the file is opened, recommending that the file be opened as read-only.
    - `CreateBackup As Variant` (optional): True to create a backup file.
    - `AddToMru As Variant` (optional): True to add this workbook to the list of recently used files. The default value is False.
    - `TextCodepage As Variant` (optional): Not used in U.S. English Microsoft Excel.
    - `TextVisualLayout As Variant` (optional): Not used in U.S. English Microsoft Excel.
    - `Local As Variant` (optional): True saves files against the language of Microsoft Excel (including control panel settings). False (default) saves files against the language of Visual Basic for Applications (VBA) (which is typically US English unless the VBA project where Workbooks.Open is run from is an old internationalized XL5/95 VBA project).

## Events (11)

- `Activate()`  
  Occurs when a workbook, worksheet, chart sheet, or embedded chart is activated.
- `Deactivate()`  
  Occurs when the chart, worksheet, or workbook is deactivated.
- `Resize()`  
  Occurs when the chart is resized.
- `MouseDown(Button As Long, Shift As Long, x As Long, y As Long)`  
  Occurs when a mouse button is pressed while the pointer is over a chart.
    - `Button As Long` (required): The mouse button that was released. Can be one of the following XlMouseButton constants: xlNoButton, xlPrimaryButton, or xlSecondaryButton.
    - `Shift As Long` (required): The state of the Shift, Ctrl, and AlShift, Ctrl, and AlttShift, Ctrl, and Alt keys when the event occurred. Can be one of or a sum of values.
    - `x As Long` (required): The _x_ coordinate of the mouse pointer in chart object client coordinates.
    - `y As Long` (required): The _y_ coordinate of the mouse pointer in chart object client coordinates.
- `MouseUp(Button As Long, Shift As Long, x As Long, y As Long)`  
  Occurs when a mouse button is released while the pointer is over a chart.
    - `Button As Long` (required): The mouse button that was released. Can be one of the following XlMouseButton constants: xlNoButton, xlPrimaryButton, or xlSecondaryButton.
    - `Shift As Long` (required): The state of the Shift, Ctrl, and Alt keys when the event occurred. Can be one of or a sum of values.
    - `x As Long` (required): The _x_ coordinate of the mouse pointer in chart object client coordinates.
    - `y As Long` (required): The _y_ coordinate of the mouse pointer in chart object client coordinates.
- `MouseMove(Button As Long, Shift As Long, x As Long, y As Long)`  
  Occurs when the position of the mouse pointer changes over a chart.
    - `Button As Long` (required): The mouse button that was released. Can be one of the following XlMouseButton constants: xlNoButton, xlPrimaryButton, or xlSecondaryButton.
    - `Shift As Long` (required): The state of the Shift, Ctrl, and Alt keys when the event occurred. Can be one of or a sum of values.
    - `x As Long` (required): The _x_ coordinate of the mouse pointer in chart object client coordinates.
    - `y As Long` (required): The _y_ coordinate of the mouse pointer in chart object client coordinates.
- `BeforeRightClick(Cancel As Boolean)`  
  Occurs when a chart element is right-clicked, before the default right-click action.
    - `Cancel As Boolean` (required): False when the event occurs. If the event procedure sets this argument to True, the default right-click action isn't performed when the procedure is finished.
- `BeforeDoubleClick(ElementID As Long, Arg1 As Long, Arg2 As Long, Cancel As Boolean)`  
  Occurs when a chart element is double-clicked, before the default double-click action.
    - `ElementID As Long` (required): The double-clicked object. The value of this parameter determines the expected values of _Arg1_ and _Arg2_. For more information about this parameter, see the Remarks section.
    - `Arg1 As Long` (required): Additional event information, depending on the value of _ElementID_. For more information about this parameter, see the Remarks section.
    - `Arg2 As Long` (required): Additional event information, depending on the value of _ElementID_. For more information about this parameter, see the Remarks section.
    - `Cancel As Boolean` (required): False when the event occurs. If the event procedure sets this argument to True, the default double-click action isn't performed when the procedure is finished.
- `Select(ElementID As Long, Arg1 As Long, Arg2 As Long)`  
  Occurs when a chart element is selected.
    - `ElementID As Long` (required): The selected chart element. For more information about this argument, see the BeforeDoubleClick event.
    - `Arg1 As Long` (required): The selected chart element. For more information about this argument, see the BeforeDoubleClick event.
    - `Arg2 As Long` (required): The selected chart element. For more information about this argument, see the BeforeDoubleClick event.
- `SeriesChange(SeriesIndex As Long, PointIndex As Long)`  
  Occurs when the user changes the value of a chart data point by choosing a bar in the chart and dragging the top edge up or down thus changing the value of the data point.
    - `SeriesIndex As Long` (required): The offset within the Series collection for the changed series.
    - `PointIndex As Long` (required): The offset within the Points collection for the changed point.
- `Calculate()`  
  Occurs after the chart plots new or changed data for the Chart object.
