# XlSlicerCrossFilterType

**Type:** Enumeration  
**Library:** Microsoft Excel 16.0 Object Library  

Specifies the type of cross filtering used by the specified slicer cache and how it is visualized.

## Constants (4)

- `xlSlicerNoCrossFilter` = 1  
  Cross filtering is turned off entirely, so all tiles are displayed and active (not dimmed) regardless of filtering selections in other slicers.
- `xlSlicerCrossFilterShowItemsWithDataAtTop` = 2  
  Cross filtering is turned on for this slicer cache, any tile with no data for a filtering selection in other slicers connected to the same data source will be dimmed. Additionally, tiles with data are moved to the top in the slicer. (Default)
- `xlSlicerCrossFilterShowItemsWithNoData` = 3  
  Cross filtering is turned on for this slicer cache, any tile with no data for a filtering selection in other slicers connected to the same data source will be dimmed.
- `xlSlicerCrossFilterHideButtonsWithNoData` = 4  
  Cross filtering is turned on for this slicer cache, any tile with no data for a filtering selection in other slicers connected to the same data source will be dimmed. Additionally, buttons will be hidden.
