# Chart

**Type:** Class  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {3B06E997-E47C-11CD-8701-00AA003F0F07}  

A customizable visualization of data that can be included in a report.

## Properties (130)

- `Application As Application  (read-only)`
- `Parent As Object  (read-only)`
- `Properties As Properties  (read-only)`
- `EventProcPrefix As String  (read/write)`
- `ControlType As Byte  (read/write)`
- `RowSource As String  (read/write)`  
  Returns or sets the name of the table/query that supplies data for the chart. Read/write String.
- `LinkChildFields As String  (read/write)`
- `LinkMasterFields As String  (read/write)`
- `Visible As Boolean  (read/write)`
- `DisplayWhen As Byte  (read/write)`
- `Enabled As Boolean  (read/write)`
- `StatusBarText As String  (read/write)`
- `TabStop As Boolean  (read/write)`
- `TabIndex As Integer  (read/write)`
- `Left As Integer  (read/write)`
- `Top As Integer  (read/write)`
- `Width As Integer  (read/write)`
- `Height As Integer  (read/write)`
- `BackStyle As Byte  (read/write)`
- `BackColor As Long  (read/write)`
- `SpecialEffect As Byte  (read/write)`
- `BorderStyle As Byte  (read/write)`
- `OldBorderStyle As Byte  (read/write)`
- `BorderColor As Long  (read/write)`
- `BorderWidth As Byte  (read/write)`
- `ShortcutMenuBar As String  (read/write)`
- `ControlTipText As String  (read/write)`
- `HelpContextId As Long  (read/write)`
- `Section As Integer  (read/write)`
- `Tag As String  (read/write)`
- `IsVisible As Boolean  (read/write)`
- `InSelection As Boolean  (read/write)`
- `OnEnter As String  (read/write)`
- `OnExit As String  (read/write)`
- `OnGotFocus As String  (read/write)`
- `OnLostFocus As String  (read/write)`
- `OnClick As String  (read/write)`
- `OnDblClick As String  (read/write)`
- `OnMouseDown As String  (read/write)`
- `OnMouseMove As String  (read/write)`
- `OnMouseUp As String  (read/write)`
- `OnKeyDown As String  (read/write)`
- `OnKeyUp As String  (read/write)`
- `OnKeyPress As String  (read/write)`
- `Name As String  (read/write)`
- `Layout As AcLayoutType  (read-only)`
- `LeftPadding As Integer  (read/write)`
- `TopPadding As Integer  (read/write)`
- `RightPadding As Integer  (read/write)`
- `BottomPadding As Integer  (read/write)`
- `GridlineStyleLeft As Byte  (read/write)`
- `GridlineStyleTop As Byte  (read/write)`
- `GridlineStyleRight As Byte  (read/write)`
- `GridlineStyleBottom As Byte  (read/write)`
- `GridlineWidthLeft As Byte  (read/write)`
- `GridlineWidthTop As Byte  (read/write)`
- `GridlineWidthRight As Byte  (read/write)`
- `GridlineWidthBottom As Byte  (read/write)`
- `GridlineColor As Long  (read/write)`
- `HorizontalAnchor As AcHorizontalAnchor  (read/write)`
- `VerticalAnchor As AcVerticalAnchor  (read/write)`
- `LayoutID As Long  (read-only)`
- `BackThemeColorIndex As Long  (read/write)`
- `BackTint As Single  (read/write)`
- `BackShade As Single  (read/write)`
- `BorderThemeColorIndex As Long  (read/write)`
- `BorderTint As Single  (read/write)`
- `BorderShade As Single  (read/write)`
- `GridlineThemeColorIndex As Long  (read/write)`
- `GridlineTint As Single  (read/write)`
- `GridlineShade As Single  (read/write)`
- `ChartAxis As String  (read/write)`  
  Returns or sets the semicolon-separated field name(s) that are used by the chart axis. Read/write String.
- `ChartLegend As String  (read/write)`  
  Returns or sets the field name that is used by the chart legend. Read/write String.
- `ChartValues As String  (read/write)`  
  Returns or sets the semicolon-separated list of field(s) used to determine the data series plotted on the value axis. Read/write String.
- `HasLegend As Boolean  (read/write)`  
  True if the chart has a legend. Read/write Boolean.
- `HasTitle As Boolean  (read/write)`  
  True if the title is visible for the specified chart. Read/write Boolean.
- `ChartTitle As String  (read/write)`  
  Returns or sets the title for the specified chart. Read/write String.
- `HasAxisTitles As Boolean  (read/write)`  
  True if the axis titles are visible for the specified chart. Read/write Boolean.
- `CategoryAxisTitle As String  (read/write)`  
  Returns or sets the title for the category axis. Read/write String.
- `PrimaryValuesAxisTitle As String  (read/write)`  
  Returns or sets the title for the primary values axis. Read/write String.
- `SecondaryValuesAxisTitle As String  (read/write)`  
  Returns or sets the title for the secondary values axis. Read/write String.
- `TransformedRowSource As String  (read-only)`  
  Returns the auto-generated SQL string that is used to support aggregation, pivoting, and ordering of chart data. Read-only String.
- `ChartType As AcChartType  (read/write)`  
  Returns or sets the chart type. Read/write AcChartType.
- `LegendPosition As AcLegendPosition  (read/write)`  
  Returns or sets the position of the legend for the specified chart. Read/write AcLegendPosition.
- `PrimaryValuesAxisMinimum As Single  (read/write)`  
  Returns or sets the minimum value that can be represented on the primary values axis. Read/write Single.
- `PrimaryValuesAxisMaximum As Single  (read/write)`  
  Returns or sets the maximum value that can be represented on the primary values axis. Read/write Single.
- `SecondaryValuesAxisMinimum As Single  (read/write)`  
  Returns or sets the minimum value that can be represented on the secondary values axis. Read/write Single.
- `SecondaryValuesAxisMaximum As Single  (read/write)`  
  Returns or sets the maximum value that can be represented on the secondary values axis. Read/write Single.
- `PrimaryValuesAxisRange As AcAxisRange  (read/write)`  
  Returns or sets the behavior for representing minimum and maximum values on the primary values axis. Read/write AcAxisRange.
- `SecondaryValuesAxisRange As AcAxisRange  (read/write)`  
  Returns or sets the behavior for representing minimum and maximum values on the secondary values axis. Read/write AcAxisRange.
- `PrimaryValuesAxisFontSize As Integer  (read/write)`  
  Returns or sets the font size used by the primary values axis. Read/write Integer.
- `PrimaryValuesAxisFontColor As Long  (read/write)`  
  Returns or sets the font color used by the primary values axis. Read/write Long.
- `PrimaryValuesAxisThemeColorIndex As Long  (read/write)`
- `PrimaryValuesAxisFontTint As Single  (read/write)`
- `PrimaryValuesAxisFontShade As Single  (read/write)`
- `SecondaryValuesAxisFontSize As Integer  (read/write)`  
  Returns or sets the font size used by the secondary values axis. Read/write Integer.
- `SecondaryValuesAxisThemeColorIndex As Long  (read/write)`
- `SecondaryValuesAxisFontTint As Single  (read/write)`
- `SecondaryValuesAxisFontShade As Single  (read/write)`
- `SecondaryValuesAxisFontColor As Long  (read/write)`  
  Returns or sets the font color used by the secondary values axis. Read/write Long.
- `CategoryAxisFontSize As Integer  (read/write)`  
  Returns or sets the font size used by the category axis. Read/write Integer.
- `CategoryAxisFontColor As Long  (read/write)`  
  Returns or sets the font color used by the category axis. Read/write Long.
- `CategoryAxisThemeColorIndex As Long  (read/write)`
- `CategoryAxisFontTint As Single  (read/write)`
- `CategoryAxisFontShade As Single  (read/write)`
- `ChartTitleFontColor As Long  (read/write)`  
  Returns or sets the font color used by the chart subtitle. Read/write Long.
- `ChartTitleThemeColorIndex As Long  (read/write)`
- `ChartTitleFontTint As Single  (read/write)`
- `ChartTitleFontShade As Single  (read/write)`
- `ChartTitleFontSize As Integer  (read/write)`  
  Returns or sets the font size used by the chart title. Read/write Integer.
- `ChartSeriesCollection As ChartSeriesCollection  (read-only)`  
  A collection of all the ChartSeries objects in the specified chart.
- `ChartAxisCollection As ChartAxisCollection  (read-only)`  
  A collection of all the ChartAxis objects in the specified chart.
- `ChartValuesCollection As ChartValuesCollection  (read-only)`  
  A collection of all the ChartValues objects in the specified chart.
- `PrimaryValuesAxisFormat As String  (read/write)`  
  Returns or sets the format of the values on the primary values axis. Read/write String.
- `SecondaryValuesAxisFormat As String  (read/write)`  
  Returns or sets the format of the values on the secondary values axis. Read/write String.
- `PrimaryValuesAxisDisplayUnits As AcAxisUnits  (read/write)`  
  Returns or sets the unit of measurement applied to the primary values axis. Read/write AcAxisUnits.
- `SecondaryValuesAxisDisplayUnits As AcAxisUnits  (read/write)`  
  Returns or sets the unit of measurement applied to the secondary values axis. Read/write AcAxisUnits.
- `ChartTitleFontName As String  (read/write)`  
  Returns or sets the name of the font used by the chart title. Read/write String.
- `LegendTextFontSize As Integer  (read/write)`  
  Returns or sets the font size used by the chart legend. Read/write Integer.
- `LegendTextFontColor As Long  (read/write)`  
  Returns or sets the font color used by the chart legend. Read/write Long.
- `LegendTextThemeColorIndex As Long  (read/write)`
- `LegendTextFontTint As Single  (read/write)`
- `LegendTextFontShade As Single  (read/write)`
- `HasSubtitle As Boolean  (read/write)`  
  True if the subtitle is visible for the specified chart. Read/write Boolean.
- `ChartSubtitle As String  (read/write)`  
  Returns or sets the subtitle for the specified chart. Read/write String.
- `ChartSubtitleFontSize As Integer  (read/write)`  
  Returns or sets the font size used by the chart subtitle. Read/write Integer.
- `ChartSubtitleFontColor As Long  (read/write)`  
  Returns or sets the font color used by the chart subtitle. Read/write Long.
- `ChartSubtitleThemeColorIndex As Long  (read/write)`
- `ChartSubtitleFontTint As Single  (read/write)`
- `ChartSubtitleFontShade As Single  (read/write)`

## Methods (4)

- `SizeToFit()`
- `Requery()`
- `SetFocus()`
- `Move(Left As Variant, [Top As Variant], [Width As Variant], [Height As Variant])`

## Events (12)

- `Enter()`
- `Exit(Cancel As Integer)`
- `GotFocus()`
- `LostFocus()`
- `Click()`
- `DblClick(Cancel As Integer)`
- `MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)`
- `MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)`
- `MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)`
- `KeyDown(KeyCode As Integer, Shift As Integer)`
- `KeyPress(KeyAscii As Integer)`
- `KeyUp(KeyCode As Integer, Shift As Integer)`
