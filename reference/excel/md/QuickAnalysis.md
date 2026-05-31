# QuickAnalysis

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244D0-0000-0000-C000-000000000046}  

Enables single-click access to data analysis features such as formulas, conditional formatting, sparklines, tables, charts, and PivotTables.

**Example:**

```vba
ActiveWorksheet.QuickAnalysis.Hide(1)
```

## Properties (3)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Excel application. Read-only.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified QuickAnalysis object. Read-only.

## Methods (2)

- `Show([XlQuickAnalysisMode As XlQuickAnalysisMode])`  
  Displays specific members of the Analysis Lens user interface.
    - `XlQuickAnalysisMode As XlQuickAnalysisMode` (optional): Indicates for which top level button the callout user interface is displayed. Can be one of the XlQuickAnalysisMode constants.
- `Hide([XlQuickAnalysisMode As XlQuickAnalysisMode])`  
  Hides specific members of the Analysis Lens user interface.
    - `XlQuickAnalysisMode As XlQuickAnalysisMode` (optional): Indicates for which top level button the callout user interface is displayed. Can be one of the XlQuickAnalysisMode constants.
