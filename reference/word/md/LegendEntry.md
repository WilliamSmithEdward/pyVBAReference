# LegendEntry

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {C4A02049-024C-4273-8934-E48CC21479A9}  

Represents a legend entry in a chart legend.

**Remarks:** The LegendEntry object is a member of the LegendEntries collection. The LegendEntries collection contains all the LegendEntry objects in the legend. Each legend entry has two parts: - The text of the entry, which is the name of the series or trendline associated with the legend entry. - The entry marker, which visually links the legend entry with its associated series or trendline in the chart. The formatting properties for the entry marker and its associated series or trendline are contained in the LegendKey object. The text of a legend entry cannot be changed. LegendEntry objects support font formatting, and they can be deleted. No pattern formatting is supported for legend entries. The position and size of entries is fixed. There is no direct way to return the series or trendline that corresponds to the legend entry. After legend entries have been deleted, the only way to restore them is to remove and re-create the legend that contained them by setting the HasLegend property for the chart to False and then back to True.

**Example:**

```vba
With ActiveDocument.InlineShapes(1)
 If .HasChart Then
 .Chart.Legend.LegendEntries(1).Font.Italic = True
 End If
End With
```

## Properties (11)

- `Parent As Object  (read-only)`  
  Returns the parent for the specified object. Read-only Object.
- `Font As ChartFont  (read-only)`  
  Returns the font of the specified object. Read-only ChartFont.
- `Index As Long  (read-only)`  
  Returns the index number of the object within the collection of similar objects. Read-only Long.
- `LegendKey As LegendKey  (read-only)`  
  Returns the legend key that is associated with the entry. Read-only LegendKey.
- `Left As Double  (read-only)`  
  Returns the distance, in points, from the left edge of the object to the left edge of the chart area. Read-only Double.
- `Top As Double  (read-only)`  
  Returns the distance, in points, from the top edge of the object to the top of the first row (on a worksheet) or the top of the chart area (on a chart). Read-only Double.
- `Width As Double  (read-only)`  
  Returns the width, in points, of the object. Read-only Double.
- `Height As Double  (read-only)`  
  Returns the height, in points, of the object. Read-only Double.
- `Format As ChartFormat  (read-only)`  
  Returns the line, fill, and effect formatting for the object. Read-only ChartFormat.
- `Application As Object  (read-only)`  
  When used without an object qualifier, returns an Application object that represents the Microsoft Word application. When used with an object qualifier, returns an Application object that represents the creator of the specified object (you can use this property with an Automation object to return the application of that object). Read-only.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.

## Methods (2)

- `Delete() As Variant`  
  Deletes the object.
- `Select() As Variant`  
  Selects the object.
