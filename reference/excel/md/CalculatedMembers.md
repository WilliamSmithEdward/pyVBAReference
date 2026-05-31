# CalculatedMembers

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024454-0000-0000-C000-000000000046}  

A collection of all the CalculatedMember objects on the specified PivotTable.

**Remarks:** Each CalculatedMember object represents a calculated member or calculated measure. Use the CalculatedMembers property of the PivotTable object to return a CalculatedMembers collection. There are three supported types of calculated members: Named Sets, Calculated Members, and Calculated Measures. Object model support has been available for all three types since Excel 2010. User interface support was made available for Named Sets in Excel 2010. In Excel 2013, the OLAP Calculated Members and Calculated Measures feature was created to build a user interface for the calculated members and measures object model. Named Sets is used exactly the same as in Excel 2010. Named Sets should continue to use the Add method, and the type XlCalculatedMemberType enumeration. Calculated Members has the following changes for Excel 2013: - It now uses the AddCalculatedMember method. - It supports the following properties of the CalculatedMember object: - ParentHierarchy property - ParentMember property - NumberFormat property Calculated Measures has the following changes for Excel 2013: - It now uses the AddCalculatedMember method. - It now uses the type XlCalculatedMemberType enumeration.

**Example:**

```vba
Sub UseCalculatedMember()
 Dim pvtTable As PivotTable
 Set pvtTable = ActiveSheet.PivotTables(1)
 pvtTable.CalculatedMembers.Add Name:="[Beef]", _
 Formula:="'{[Product].[All Products].Children}'", _
 Type:=xlCalculatedSet

End Sub
```

## Properties (7)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Count As Long  (read-only)`  
  Returns a Long value that represents the number of objects in the collection.
- `Item As CalculatedMember  (read-only)`  
  Returns a single object from a collection.
- `_Default As CalculatedMember  (read-only)`
- `_NewEnum As IUnknown  (read-only)`

## Methods (2)

- `Add(Name As String, Formula As Variant, [SolveOrder As Variant], [Type As Variant], [Dynamic As Variant], [DisplayFolder As Variant], [HierarchizeDistinct As Variant]) As CalculatedMember`  
  Adds a calculated field or calculated item to a PivotTable. Returns a CalculatedMember object.
    - `Name As String` (required): The name of the calculated member.
    - `Formula As Variant` (required): The formula of the calculated member.
    - `SolveOrder As Variant` (optional): The solve order for the calculated member.
    - `Type As Variant` (optional): The type of calculated member.
    - `Dynamic As Variant` (optional): Specifies if the calculated member is recalculated with every update.
    - `DisplayFolder As Variant` (optional): The name of the display folder for the calculated member.
    - `HierarchizeDistinct As Variant` (optional): Specifies whether to order and remove duplicates when displaying the hierarchy of the calculated member in a PivotTable report based on an OLAP cube.
- `AddCalculatedMember(Name As String, Formula As Variant, [SolveOrder As Variant], [Type As Variant], [DisplayFolder As Variant], [MeasureGroup As Variant], [ParentHierarchy As Variant], [ParentMember As Variant], [NumberFormat As Variant]) As CalculatedMember`  
  Adds a calculated field or calculated item to a PivotTable.
    - `Name As String` (required): The name of the calculated member.
    - `Formula As Variant` (required): The formula of the calculated member.
    - `SolveOrder As Variant` (optional): The solve order for the calculated member.
    - `Type As Variant` (optional): The type of calculated member.
    - `DisplayFolder As Variant` (optional): A folder that exists to display calculated measures.
    - `MeasureGroup As Variant` (optional): The group to which the calculated member belongs.
    - `ParentHierarchy As Variant` (optional): The parent path of the ParentMember.
    - `ParentMember As Variant` (optional): The parent of the calculated member.
    - `NumberFormat As Variant` (optional): The format of numbers used for calculated members.
