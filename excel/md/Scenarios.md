# Scenarios

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00020896-0000-0000-C000-000000000046}  

A collection of all the Scenario objects on the specified worksheet.

**Remarks:** A scenario is a group of input values (called _changing cells_) that's named and saved.

**Example:**

```vba
Worksheets("options").Scenarios.CreateSummary _
 resultCells:=Worksheets("options").Range("j10,j20")
```

## Properties (4)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Count As Long  (read-only)`  
  Returns a Long value that represents the number of objects in the collection.

## Methods (5)

- `Add(Name As String, ChangingCells As Variant, [Values As Variant], [Comment As Variant], [Locked As Variant], [Hidden As Variant]) As Scenario`  
  Creates a new scenario and adds it to the list of scenarios that are available for the current worksheet.
    - `Name As String` (required): The scenario name.
    - `ChangingCells As Variant` (required): A Range object that refers to the changing cells for the scenario.
    - `Values As Variant` (optional): An array that contains the scenario values for the cells in _ChangingCells_. If this argument is omitted, the scenario values are assumed to be the current values in the cells in _ChangingCells_.
    - `Comment As Variant` (optional): A string that specifies comment text for the scenario. When a new scenario is added, the author's name and date are automatically added at the beginning of the comment text.
    - `Locked As Variant` (optional): True to lock the scenario to prevent changes. The default value is True.
    - `Hidden As Variant` (optional): True to hide the scenario. The default value is False.
- `CreateSummary([ReportType As XlSummaryReportType], [ResultCells As Variant]) As Variant`  
  Creates a new worksheet that contains a summary report for the scenarios on the specified worksheet. Variant.
    - `ReportType As XlSummaryReportType` (optional): Specifies whether the summary report is a PivotTable or a standard summary.
    - `ResultCells As Variant` (optional): A Range object that represents the result cells on the specified worksheet. Normally, this range refers to one or more cells containing the formulas that depend on the changing cell values for your model; that is, the cells that show the results of a particular scenario. If this argument is omitted, there are no result cells included in the report.
- `Item(Index As Variant) As Scenario`  
  Returns a single object from a collection.
    - `Index As Variant` (required): The name or index number for the object.
- `Merge(Source As Variant) As Variant`  
  Merges the scenarios from another sheet into the Scenarios collection.
    - `Source As Variant` (required): The name of the sheet that contains scenarios to be merged, or a Worksheet object that represents that sheet.
- `_NewEnum() As IUnknown`
