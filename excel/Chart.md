# Chart

**Type:** Class  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00020821-0000-0000-C000-000000000046}  

## Properties (61)

- `Application As Application  (read-only)`
- `Creator As XlCreator  (read-only)`
- `Parent As Object  (read-only)`
- `CodeName As String  (read-only)`
- `_CodeName As String  (read/write)`
- `Index As Long  (read-only)`
- `Name As String  (read/write)`
- `Next As Object  (read-only)`
- `PageSetup As PageSetup  (read-only)`
- `Previous As Object  (read-only)`
- `ProtectContents As Boolean  (read-only)`
- `ProtectDrawingObjects As Boolean  (read-only)`
- `ProtectionMode As Boolean  (read-only)`
- `Visible As XlSheetVisibility  (read/write)`
- `Shapes As Shapes  (read-only)`
- `AutoScaling As Boolean  (read/write)`
- `ChartArea As ChartArea  (read-only)`
- `ChartTitle As ChartTitle  (read-only)`
- `DataTable As DataTable  (read-only)`
- `DepthPercent As Long  (read/write)`
- `DisplayBlanksAs As XlDisplayBlanksAs  (read/write)`
- `Elevation As Long  (read/write)`
- `Floor As Floor  (read-only)`
- `GapDepth As Long  (read/write)`
- `HasAxis As Variant  (read/write)`
- `HasDataTable As Boolean  (read/write)`
- `HasLegend As Boolean  (read/write)`
- `HasTitle As Boolean  (read/write)`
- `HeightPercent As Long  (read/write)`
- `Hyperlinks As Hyperlinks  (read-only)`
- `Legend As Legend  (read-only)`
- `Perspective As Long  (read/write)`
- `PlotArea As PlotArea  (read-only)`
- `PlotVisibleOnly As Boolean  (read/write)`
- `RightAngleAxes As Variant  (read/write)`
- `Rotation As Variant  (read/write)`
- `ChartType As XlChartType  (read/write)`
- `Walls As Walls  (read-only)`
- `BarShape As XlBarShape  (read/write)`
- `PlotBy As XlRowCol  (read/write)`
- `ProtectFormatting As Boolean  (read/write)`
- `ProtectData As Boolean  (read/write)`
- `ProtectSelection As Boolean  (read/write)`
- `PivotLayout As PivotLayout  (read-only)`
- `Tab As Tab  (read-only)`
- `MailEnvelope As MsoEnvelope  (read-only)`
- `ShowDataLabelsOverMaximum As Boolean  (read/write)`
- `SideWall As Walls  (read-only)`
- `BackWall As Walls  (read-only)`
- `ChartStyle As Variant  (read/write)`
- `PrintedCommentPages As Long  (read-only)`
- `ShowReportFilterFieldButtons As Boolean  (read/write)`
- `ShowLegendFieldButtons As Boolean  (read/write)`
- `ShowAxisFieldButtons As Boolean  (read/write)`
- `ShowValueFieldButtons As Boolean  (read/write)`
- `ShowAllFieldButtons As Boolean  (read/write)`
- `CategoryLabelLevel As XlCategoryLabelLevel  (read/write)`
- `SeriesNameLevel As XlSeriesNameLevel  (read/write)`
- `ChartColor As Variant  (read/write)`
- `ShowExpandCollapseEntireFieldButtons As Boolean  (read/write)`
- `DisplayValueNotAvailableAsBlank As Boolean  (read/write)`

## Methods (39)

- `Activate()`
- `Copy([Before As Variant], [After As Variant])`
- `Delete()`
- `Move([Before As Variant], [After As Variant])`
- `PrintPreview([EnableChanges As Variant])`
- `Select([Replace As Variant])`
- `Unprotect([Password As Variant])`
- `Axes([Type As Variant], [AxisGroup As XlAxisGroup]) As Object`
- `SetBackgroundPicture(Filename As String)`
- `ChartGroups([Index As Variant]) As Object`
- `ChartObjects([Index As Variant]) As Object`
- `ChartWizard([Source As Variant], [Gallery As Variant], [Format As Variant], [PlotBy As Variant], [CategoryLabels As Variant], [SeriesLabels As Variant], [HasLegend As Variant], [Title As Variant], [CategoryTitle As Variant], [ValueTitle As Variant], [ExtraTitle As Variant])`
- `CheckSpelling([CustomDictionary As Variant], [IgnoreUppercase As Variant], [AlwaysSuggest As Variant], [SpellLang As Variant])`
- `CopyPicture([Appearance As XlPictureAppearance], [Format As XlCopyPictureFormat], [Size As XlPictureAppearance])`
- `Evaluate(Name As Variant) As Variant`
- `_Evaluate(Name As Variant) As Variant`
- `Location(Where As XlChartLocation, [Name As Variant]) As Chart`
- `OLEObjects([Index As Variant]) As Object`
- `Paste([Type As Variant])`
- `SeriesCollection([Index As Variant]) As Object`
- `GetChartElement(x As Long, y As Long, ElementID As Long, Arg1 As Long, Arg2 As Long)`
- `SetSourceData(Source As Range, [PlotBy As Variant])`
- `Export(Filename As String, [FilterName As Variant], [Interactive As Variant]) As Boolean`
- `Refresh()`
- `ApplyDataLabels([Type As XlDataLabelsType], [LegendKey As Variant], [AutoText As Variant], [HasLeaderLines As Variant], [ShowSeriesName As Variant], [ShowCategoryName As Variant], [ShowValue As Variant], [ShowPercentage As Variant], [ShowBubbleSize As Variant], [Separator As Variant])`
- `Protect([Password As Variant], [DrawingObjects As Variant], [Contents As Variant], [Scenarios As Variant], [UserInterfaceOnly As Variant])`
- `ApplyLayout(Layout As Long, [ChartType As Variant])`
- `SetElement(Element As MsoChartElementType)`
- `PrintOut([From As Variant], [To As Variant], [Copies As Variant], [Preview As Variant], [ActivePrinter As Variant], [PrintToFile As Variant], [Collate As Variant], [PrToFileName As Variant])`
- `ApplyChartTemplate(Filename As String)`
- `SaveChartTemplate(Filename As String)`
- `SetDefaultChart(Name As Variant)`
- `ClearToMatchStyle()`
- `FullSeriesCollection([Index As Variant]) As Object`
- `ClearToMatchColorStyle()`
- `ExportAsFixedFormat(Type As XlFixedFormatType, [Filename As Variant], [Quality As Variant], [IncludeDocProperties As Variant], [IgnorePrintAreas As Variant], [From As Variant], [To As Variant], [OpenAfterPublish As Variant], [FixedFormatExtClassPtr As Variant], [WorkIdentity As Variant])`
- `SetProperty(ID As String, Value As Variant)`
- `GetProperty(ID As String) As Variant`
- `SaveAs(Filename As String, [FileFormat As Variant], [Password As Variant], [WriteResPassword As Variant], [ReadOnlyRecommended As Variant], [CreateBackup As Variant], [AddToMru As Variant], [TextCodepage As Variant], [TextVisualLayout As Variant], [Local As Variant])`

## Events (11)

- `Activate()`
- `Deactivate()`
- `Resize()`
- `MouseDown(Button As Long, Shift As Long, x As Long, y As Long)`
- `MouseUp(Button As Long, Shift As Long, x As Long, y As Long)`
- `MouseMove(Button As Long, Shift As Long, x As Long, y As Long)`
- `BeforeRightClick(Cancel As Boolean)`
- `BeforeDoubleClick(ElementID As Long, Arg1 As Long, Arg2 As Long, Cancel As Boolean)`
- `Select(ElementID As Long, Arg1 As Long, Arg2 As Long)`
- `SeriesChange(SeriesIndex As Long, PointIndex As Long)`
- `Calculate()`
