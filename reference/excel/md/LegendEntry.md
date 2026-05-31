# LegendEntry

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000208BA-0000-0000-C000-000000000046}  

Represents a legend entry in a chart legend.

**Remarks:** The LegendEntry object is a member of the LegendEntries collection. The LegendEntries collection contains all the LegendEntry objects in the legend. Each legend entry has two parts: the text of the entry, which is the name of the series associated with the legend entry; and an entry marker, which visually links the legend entry with its associated series or trendline in the chart. Formatting properties for the entry marker and its associated series or trendline are contained in the LegendKey object. The text of a legend entry cannot be changed. LegendEntry objects support font formatting, and they can be deleted. No pattern formatting is supported for legend entries. The position and size of entries is fixed. There's no direct way to return the series or trendline corresponding to the legend entry. After legend entries have been deleted, the only way to restore them is to remove and recreate the legend that contained them by setting the HasLegend property for the chart to False and then back to True.

**Example:**

```vba
Worksheets("sheet1").ChartObjects(1).Chart _
 .Legend.LegendEntries(1).Font.Italic = True
```

## Properties (11)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Font As Font  (read-only)`  
  Returns a Font object that represents the font of the specified object.
- `Index As Long  (read-only)`  
  Returns a Long value that represents the index number of the object within the collection of similar objects.
- `LegendKey As LegendKey  (read-only)`  
  Returns a LegendKey object that represents the legend key associated with the entry.
- `Left As Double  (read-only)`  
  Returns a Double value that represents the distance, in points, from the left edge of the object to the left edge of the chart area.
- `Top As Double  (read-only)`  
  Returns a Double value that represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).
- `Width As Double  (read-only)`  
  Returns a Double value that represents the width, in points, of the object.
- `Height As Double  (read-only)`  
  Returns a Double value that represents the height, in points, of the object.
- `Format As ChartFormat  (read-only)`  
  Returns the ChartFormat object. Read-only.

## Methods (2)

- `Delete() As Variant`  
  Deletes the object.
- `Select() As Variant`  
  Selects the object.
