# FullSeriesCollection

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244DC-0000-0000-C000-000000000046}  

Represents the full set of Series objects in a chart.

**Remarks:** The FullSeriesCollection object enables you to get a filtered out Series object and filter it back in. It also enables you to iterate over the full set of Series objects, filtered out or visible, programmatically. By having the existing SeriesCollection object contain only the visible series, you can programmatically perform operations on only the visible series. It also prevents Microsoft Excel from breaking existing chart solutions on charts with filtered out data.

**Example:**

```vba
MsgBox Chart(1).FullSeriesCollection.Item(2).Name
```

## Properties (4)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Excel application. Read-only.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified FullSeriesCollection object. Read-only.
- `Count As Long  (read-only)`  
  Returns a Long value that represents the number of Series objects in the FullSeriesCollection object. Read-only.

## Methods (3)

- `Item(Index As Variant) As Series`  
  Returns a single object from the FullSeriesCollection object.
    - `Index As Variant` (required): The index number for the object.
- `_NewEnum() As IUnknown`
- `_Default(Index As Variant) As Series`
