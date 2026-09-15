# Borders

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {914934CA-5A91-11CF-8700-00AA0060263B}  

A collection of LineFormat objects that represent the borders and diagonal lines of a cell or range of cells in a table.

**Remarks:** Each Cell object or CellRange collection has six elements in the Borders collection. You cannot add objects to the Borders collection. Use Borders (index), where index identifies the cell border or diagonal line, to return a single Border object. The index value can be any PPBorderType constant.

**Example:**

```vba
ActiveWindow.Selection.ShapeRange.Table.Rows(2).Cells.Borders(ppBorderBottom).DashStyle = msoLineDash
```

## Properties (3)

- `Count As Long  (read-only)`  
  Returns the number of objects in the specified collection. Read-only.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.

## Methods (1)

- `Item(BorderType As PpBorderType) As LineFormat`  
  Returns a LineFormat object for the specified border from the Borders collection.
    - `BorderType As PpBorderType` (required): Specifies which border of a cell or cell range is to be returned.
