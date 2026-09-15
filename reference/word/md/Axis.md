# Axis

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {7EBC66BD-F788-42C3-91F4-E8C841A69005}  

Represents a single axis in a chart.

**Remarks:** The Axis object is a member of the Axes collection. Use Axes ( _Type_ , _AxisGroup_ ) where _Type_ is the axis type and _AxisGroup_ is the axis group to return a single Axis object. _Type_ can be one of the following XlAxisType constants: xlCategory, xlSeries, or xlValue. _AxisGroup_ can be one of the following XlAxisGroup constants: xlPrimary or xlSecondary. For more information, see the Axes method.

**Example:**

```vba
With ActiveDocument.InlineShapes(1)
 If .HasChart Then
 With .Chart.Axes(xlCategory)
 .HasTitle = True
 .AxisTitle.Caption = "1994"
 End With
 End If
End With
```

## Properties (49)

- `AxisBetweenCategories As Boolean  (read/write)`  
  True if the value axis crosses the category axis between categories. Read/write Boolean.
- `AxisGroup As XlAxisGroup  (read-only)`  
  Returns the type of axis group. Read-only XlAxisGroup.
- `AxisTitle As AxisTitle  (read-only)`  
  Returns the title of the specified axis. Read-only AxisTitle.
- `CategoryNames As Variant  (read/write)`  
  Returns or sets all the category names as a text array for the specified axis. Read/write Variant.
- `Crosses As XlAxisCrosses  (read/write)`  
  Returns or sets the point on the specified axis where the other axis crosses. Read/write Long.
- `CrossesAt As Double  (read/write)`  
  Returns or sets the point on the value axis where the category axis crosses it. Applies only to the value axis. Read/write Double.
- `HasMajorGridlines As Boolean  (read/write)`  
  True if the axis has major gridlines. Read/write Boolean.
- `HasMinorGridlines As Boolean  (read/write)`  
  True if the axis has minor gridlines. Read/write Boolean.
- `HasTitle As Boolean  (read/write)`  
  True if the axis or chart has a visible title. Read/write Boolean.
- `MajorGridlines As Gridlines  (read-only)`  
  Returns the major gridlines for the specified axis. Read-only Gridlines.
- `MajorTickMark As XlTickMark  (read/write)`  
  Returns or sets the type of major tick mark for the specified axis. Read/write XlTickMark.
- `MajorUnit As Double  (read/write)`  
  Returns or sets the major units for the value axis. Read/write Double.
- `LogBase As Double  (read/write)`  
  Returns or sets the base of the logarithm when you are using log scales. Read/write Double.
- `TickLabelSpacingIsAuto As Boolean  (read/write)`  
  Returns or sets a value that indicates whether the tick label spacing is automatic. Read/write Boolean.
- `MajorUnitIsAuto As Boolean  (read/write)`  
  True if Microsoft Word calculates the major units for the value axis. Read/write Boolean.
- `MaximumScale As Double  (read/write)`  
  Returns or sets the maximum value on the value axis. Read/write Double.
- `MaximumScaleIsAuto As Boolean  (read/write)`  
  True if Microsoft Word calculates the maximum value for the value axis. Read/write Boolean.
- `MinimumScale As Double  (read/write)`  
  Returns or sets the minimum value on the value axis. Read/write Double.
- `MinimumScaleIsAuto As Boolean  (read/write)`  
  True if Microsoft Word calculates the minimum value for the value axis. Read/write Boolean.
- `MinorGridlines As Gridlines  (read-only)`  
  Returns the minor gridlines for the specified axis. Read-only Gridlines.
- `MinorTickMark As XlTickMark  (read/write)`  
  Returns or sets the type of minor tick mark for the specified axis. Read/write XlTickMark.
- `MinorUnit As Double  (read/write)`  
  Returns or sets the minor units on the value axis. Read/write Double.
- `MinorUnitIsAuto As Boolean  (read/write)`  
  True if Microsoft Word calculates minor units for the value axis. Read/write Boolean.
- `ReversePlotOrder As Boolean  (read/write)`  
  True if Microsoft Word plots data points from last to first. Read/write Boolean.
- `ScaleType As XlScaleType  (read/write)`  
  Returns or sets the value axis scale type. Read/write XlScaleType.
- `TickLabelPosition As XlTickLabelPosition  (read/write)`  
  Describes the position of tick-mark labels on the specified axis. Read/write XlTickLabelPosition.
- `TickLabels As TickLabels  (read-only)`  
  Returns the tick-mark labels for the specified axis. Read-only TickLabels.
- `TickLabelSpacing As Long  (read/write)`  
  Returns or sets the number of categories or series between tick-mark labels. Read/write Long.
- `TickMarkSpacing As Long  (read/write)`  
  Returns or sets the number of categories or series between tick marks. Read/write Long.
- `Type As XlAxisType  (read/write)`  
  Returns the axis type. Read-only XlAxisType.
- `BaseUnit As XlTimeUnit  (read/write)`  
  Returns or sets the base unit for the specified category axis. Read/write XlTimeUnit.
- `BaseUnitIsAuto As Boolean  (read/write)`  
  True if Microsoft Word chooses appropriate base units for the specified category axis. The default is True. Read/write Boolean.
- `MajorUnitScale As XlTimeUnit  (read/write)`  
  Returns or sets the major unit scale value for the category axis when the CategoryType property is set to xlTimeScale. Read/write XlTimeUnit.
- `MinorUnitScale As XlTimeUnit  (read/write)`  
  Returns or sets the minor unit scale value for the category axis when the CategoryType property is set to xlTimeScale. Read/write XlTimeUnit.
- `CategoryType As XlCategoryType  (read/write)`  
  Returns or sets the category axis type. Read/write XlCategoryType.
- `Left As Double  (read-only)`  
  Returns the distance, in points, from the left edge of the object to the left edge of the chart area. Read-only Double.
- `Top As Double  (read-only)`  
  Returns the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart). Read-only Double.
- `Width As Double  (read-only)`  
  Returns the width, in points, of the object. Read-only Double.
- `Height As Double  (read-only)`  
  Returns the height, in points, of the object. Read-only Double.
- `DisplayUnit As XlDisplayUnit  (read/write)`  
  Returns or sets the unit label for the value axis. Read/write XlDisplayUnit, xlCustom, or xlNone.
- `DisplayUnitCustom As Double  (read/write)`  
  If the value of the DisplayUnit property is xlCustom, returns or sets the value of the displayed units. Read/write Double.
- `HasDisplayUnitLabel As Boolean  (read/write)`  
  True if the label specified by the DisplayUnit or DisplayUnitCustom property is displayed on the specified axis. The default is True. Read/write Boolean.
- `DisplayUnitLabel As DisplayUnitLabel  (read-only)`  
  Returns the DisplayUnitLabel object for the specified axis. Returns null if the HasDisplayUnitLabel property is set to False. Read-only.
- `Border As ChartBorder  (read-only)`  
  Returns the border of the object. Read-only ChartBorder.
- `Format As ChartFormat  (read-only)`  
  Returns the line, fill, and effect formatting for the object. Read-only ChartFormat.
- `Application As Object  (read-only)`  
  When used without an object qualifier, returns an Application object that represents the Microsoft Word application. When used with an object qualifier, returns an Application object that represents the creator of the specified object (you can use this property with an Automation object to return the application of that object). Read-only.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent for the specified object. Read-only Object.
- `CategorySortOrder As XlCategorySortOrder  (read/write)`

## Methods (4)

- `Delete() As Variant`  
  Deletes the object.
- `Select() As Variant`  
  Selects the object.
- `SetProperty(ID As String, Value As Variant)`
- `GetProperty(ID As String) As Variant`
