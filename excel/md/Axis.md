# Axis

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00020848-0000-0000-C000-000000000046}  

Represents a single axis in a chart.

**Remarks:** The Axis object is a member of the Axes collection. Use Axes (_type_, _group_), where _type_ is the axis type and _group_ is the axis group, to return a single Axis object. - _Type_ can be one of the following XlAxisType constants: xlCategory, xlSeriesAxis, or xlValue. - _Group_ can be one of the following XlAxisGroup constants: xlPrimary or xlSecondary. For more information, see the Axes method of the Chart object.

**Example:**

```vba
With Charts("chart1").Axes(xlCategory)
 .HasTitle = True
 .AxisTitle.Caption = "1994"
End With
```

## Properties (49)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `AxisBetweenCategories As Boolean  (read/write)`  
  True if the value axis crosses the category axis between categories. Read/write Boolean.
- `AxisGroup As XlAxisGroup  (read-only)`  
  Returns the group for the specified axis. Read-only.
- `AxisTitle As AxisTitle  (read-only)`  
  Returns an AxisTitle object that represents the title of the specified axis. Read-only.
- `Border As Border  (read-only)`  
  Returns a Border object that represents the border of the object.
- `CategoryNames As Variant  (read/write)`  
  Returns or sets all the category names for the specified axis as a text array. When you set this property, you can set it to either an array or a Range object that contains the category names. Read/write Variant.
- `Crosses As XlAxisCrosses  (read/write)`  
  Returns or sets the point on the specified axis where the other axis crosses. Read/write Long.
- `CrossesAt As Double  (read/write)`  
  Returns or sets the point on the value axis where the category axis crosses it. Applies only to the value axis. Read/write Double.
- `HasMajorGridlines As Boolean  (read/write)`  
  True if the axis has major gridlines. Only axes in the primary axis group can have gridlines. Read/write Boolean.
- `HasMinorGridlines As Boolean  (read/write)`  
  True if the axis has minor gridlines. Only axes in the primary axis group can have gridlines. Read/write Boolean.
- `HasTitle As Boolean  (read/write)`  
  True if the axis or chart has a visible title. Read/write Boolean.
- `MajorGridlines As Gridlines  (read-only)`  
  Returns a Gridlines object that represents the major gridlines for the specified axis. Only axes in the primary axis group can have gridlines. Read-only.
- `MajorTickMark As XlTickMark  (read/write)`  
  Returns or sets the type of major tick mark for the specified axis. Read/write XlTickMark.
- `MajorUnit As Double  (read/write)`  
  Returns or sets the major units for the value axis. Read/write Double.
- `MajorUnitIsAuto As Boolean  (read/write)`  
  True if Microsoft Excel calculates the major units for the value axis. Read/write Boolean.
- `MaximumScale As Double  (read/write)`  
  Returns or sets the maximum value on the value axis. Read/write Double.
- `MaximumScaleIsAuto As Boolean  (read/write)`  
  True if Microsoft Excel calculates the maximum value for the value axis. Read/write Boolean.
- `MinimumScale As Double  (read/write)`  
  Returns or sets the minimum value on the value axis. Read/write Double.
- `MinimumScaleIsAuto As Boolean  (read/write)`  
  True if Microsoft Excel calculates the minimum value for the value axis. Read/write Boolean.
- `MinorGridlines As Gridlines  (read-only)`  
  Returns a Gridlines object that represents the minor gridlines for the specified axis. Only axes in the primary axis group can have gridlines. Read-only.
- `MinorTickMark As XlTickMark  (read/write)`  
  Returns or sets the type of minor tick mark for the specified axis. Read/write XlTickMark.
- `MinorUnit As Double  (read/write)`  
  Returns or sets the minor units on the value axis. Read/write Double.
- `MinorUnitIsAuto As Boolean  (read/write)`  
  True if Microsoft Excel calculates minor units for the value axis. Read/write Boolean.
- `ReversePlotOrder As Boolean  (read/write)`  
  True if Microsoft Excel plots data points from last to first. Read/write Boolean.
- `ScaleType As XlScaleType  (read/write)`  
  Returns or sets the value axis scale type. Read/write XlScaleType.
- `TickLabelPosition As XlTickLabelPosition  (read/write)`  
  Describes the position of tick-mark labels on the specified axis. Read/write XlTickLabelPosition.
- `TickLabels As TickLabels  (read-only)`  
  Returns a TickLabels object that represents the tick-mark labels for the specified axis. Read-only.
- `TickLabelSpacing As Long  (read/write)`  
  Returns or sets the number of categories or series between tick-mark labels. Applies only to category and series axes. Can be a value from 1 through 31999. Read/write Long.
- `TickMarkSpacing As Long  (read/write)`  
  Returns or sets the number of categories or series between tick marks. Applies only to category and series axes. Can be a value from 1 through 31999. Read/write Long.
- `Type As XlAxisType  (read/write)`  
  Returns an XlAxisType value that represents the Axis type.
- `BaseUnit As XlTimeUnit  (read/write)`  
  Returns or sets the base unit for the specified category axis. Read/write XlTimeUnit.
- `BaseUnitIsAuto As Boolean  (read/write)`  
  True if Microsoft Excel chooses appropriate base units for the specified category axis. The default value is True. Read/write Boolean.
- `MajorUnitScale As XlTimeUnit  (read/write)`  
  Returns or sets the major unit scale value for the category axis when the CategoryType property is set to xlTimeScale. Read/write XlTimeUnit.
- `MinorUnitScale As XlTimeUnit  (read/write)`  
  Returns or sets the minor unit scale value for the category axis when the CategoryType property is set to xlTimeScale. Read/write XlTimeUnit.
- `CategoryType As XlCategoryType  (read/write)`  
  Returns or sets the category axis type. Read/write XlCategoryType.
- `Left As Double  (read-only)`  
  Returns a Double value that represents the distance, in points, from the left edge of the object to the left edge of the chart area.
- `Top As Double  (read-only)`  
  Returns a Double value that represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).
- `Width As Double  (read-only)`  
  Returns a Double value that represents the width, in points, of the object.
- `Height As Double  (read-only)`  
  Returns a Double value that represents the height, in points, of the object.
- `DisplayUnit As XlDisplayUnit  (read/write)`  
  Returns or sets the unit label for the value axis. Read/write XlDisplayUnit, xlCustom, or xlNone.
- `DisplayUnitCustom As Double  (read/write)`  
  If the value of the DisplayUnit property is xlCustom, the DisplayUnitCustom property returns or sets the value of the displayed units. The value must be from 0 through 10E307. Read/write Double.
- `HasDisplayUnitLabel As Boolean  (read/write)`  
  True if the label specified by the DisplayUnit or DisplayUnitCustom property is displayed on the specified axis. The default value is True. Read/write Boolean.
- `DisplayUnitLabel As DisplayUnitLabel  (read-only)`  
  Returns the DisplayUnitLabel object for the specified axis. Returns null if the HasDisplayUnitLabel property is set to False. Read-only.
- `LogBase As Double  (read/write)`  
  Returns or sets the base of the logarithm when you are using log scales. Read/write Double.
- `TickLabelSpacingIsAuto As Boolean  (read/write)`  
  Returns or sets whether or not the tick label spacing is automatic. Read/write Boolean.
- `Format As ChartFormat  (read-only)`  
  Returns the ChartFormat object. Read-only.
- `CategorySortOrder As XlCategorySortOrder  (read/write)`

## Methods (4)

- `Delete() As Variant`  
  Deletes the object.
- `Select() As Variant`  
  Selects the object.
- `SetProperty(ID As String, Value As Variant)`
- `GetProperty(ID As String) As Variant`
