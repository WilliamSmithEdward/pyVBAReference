# CalculatedMember

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024455-0000-0000-C000-000000000046}  

Represents the calculated fields, calculated items, and named sets for PivotTables with Online Analytical Processing (OLAP) data sources.

**Remarks:** Use the Add method or the Item property of the CalculatedMembers collection to return a CalculatedMember object. With a CalculatedMember object, you can check the validity of a calculated field or item in a PivotTable by using the IsValid property.

**Example:**

```vba
Sub CheckValidity()

 Dim pvtTable As PivotTable
 Dim pvtCache As PivotCache

 Set pvtTable = ActiveSheet.PivotTables(1)
 Set pvtCache = Application.ActiveWorkbook.PivotCaches.Item(1)

 ' Handle run-time error if external source is not an OLEDB data source.
 On Error GoTo Not_OLEDB

 ' Check connection setting and make connection if necessary.
 If pvtCache.IsConnected = False Then
 pvtCache.MakeConnection
 End If

 ' Check if calculated member is valid.
 If pvtTable.CalculatedMembers.Item(1).IsValid = True Then
 MsgBox "The calculated member is valid."
 Else
 MsgBox "The calculated member is not valid."
 End If

End Sub
```

## Properties (18)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Name As String  (read-only)`  
  Returns a String value that represents the name of the object.
- `Formula As String  (read-only)`  
  Returns a String value that represents the member's formula in multidimensional expressions (MDX) syntax.
- `SourceName As String  (read-only)`  
  Returns a String value that represents the specified object's name as it appears in the original source data for the specified PivotTable report.
- `SolveOrder As Long  (read-only)`  
  Returns a Long specifying the value of the calculated member's solve order MDX (Mulitdimensional Expression) argument. The default value is zero. Read-only.
- `IsValid As Boolean  (read-only)`  
  Returns a Boolean that indicates whether the specified calculated member has been successfully instantiated with the OLAP provider during the current session.
- `_Default As String  (read-only)`
- `Type As XlCalculatedMemberType  (read-only)`  
  Returns an XlCalculatedMemberType value that represents the calculated member type.
- `Dynamic As Boolean  (read-only)`  
  Returns whether the specified named set is recalculated with every update. Read-only Boolean.
- `DisplayFolder As String  (read-only)`  
  Returns the display folder name for a named set. Read-only.
- `HierarchizeDistinct As Boolean  (read/write)`  
  Returns or sets whether to order and remove duplicates when displaying the hierarchy of the specified named set in a PivotTable report based on an OLAP cube. Read/write Boolean.
- `FlattenHierarchies As Boolean  (read/write)`  
  Returns or sets whether items from all levels of the hierarchy of the specified named set are displayed in the same field of a PivotTable report based on an OLAP cube. Read/write Boolean.
- `MeasureGroup As String  (read-only)`  
  Returns the associated measure group. Read-only String.
- `ParentHierarchy As String  (read-only)`  
  Returns the name of the current parent hierarchy from the hierarchies that are available on the cube. Read-only String.
- `ParentMember As String  (read-only)`  
  Returns the name of the parent member for the parent hierarchy. Read-only String.
- `NumberFormat As XlCalcMemNumberFormatType  (read-only)`  
  Returns an XlCalcMemNumberFormatType value that represents the number format of the calculated member. The default value is xlNumberFormatTypeDefault. Read-only.

## Methods (1)

- `Delete()`  
  Deletes the object.
