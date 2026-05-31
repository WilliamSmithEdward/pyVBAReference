# Scenario

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00020897-0000-0000-C000-000000000046}  

Represents a scenario on a worksheet.

**Remarks:** A scenario is a group of input values (called _changing cells_) that's named and saved. The Scenario object is a member of the Scenarios collection. The Scenarios collection contains all the defined scenarios for a worksheet.

**Example:**

```vba
Worksheets("options").Scenarios("typical").Show
```

## Properties (10)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `ChangingCells As Range  (read-only)`  
  Returns a Range object that represents the changing cells for a scenario. Read-only.
- `Comment As String  (read/write)`  
  Returns or sets a String value that represents the comment associated with the scenario.
- `Hidden As Boolean  (read/write)`  
  Returns or sets a Boolean value that indicates if the scenario is hidden.
- `Index As Long  (read-only)`  
  Returns a Long value that represents the index number of the object within the collection of similar objects.
- `Locked As Boolean  (read/write)`  
  Returns or sets a Boolean value that indicates if the object is locked.
- `Name As String  (read/write)`  
  Returns or sets a String value representing the name of the object.
- `Values As Variant  (read-only)`  
  Returns a Variant array that contains the current values of the changing cells for the scenario.

## Methods (3)

- `ChangeScenario(ChangingCells As Variant, [Values As Variant]) As Variant`  
  Changes the scenario to have a new set of changing cells and (optionally) scenario values.
    - `ChangingCells As Variant` (required): A Range object that specifies the new set of changing cells for the scenario. The changing cells must be on the same sheet as the scenario.
    - `Values As Variant` (optional): An array that contains the new scenario values for the changing cells. If this argument is omitted, the scenario values are assumed to be the current values in the changing cells.
- `Delete() As Variant`  
  Deletes the object.
- `Show() As Variant`  
  Shows the scenario by inserting its values on the worksheet. The affected cells are the changing cells of the scenario.
