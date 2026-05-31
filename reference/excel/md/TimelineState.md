# TimelineState

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000244DF-0000-0000-C000-000000000046}  

The timeline-specific state of a SlicerCache object.

**Remarks:** Supported contiguous ranges can be set through the SetFilterDateRange method. When the timeline has such a contiguous filter state, the state can be retrieved from the two properties StartDate and EndDate. Any state that the filter may have, including non-contiguous states, can be retrieved through the three properties FilterType, FilterValue1, and FilterValue2.

## Properties (9)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Excel application. Read-only.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified TimelineState object. Read-only.
- `StartDate As Variant  (read-only)`  
  Returns the start of the filtering date range. Read-only Variant.
- `EndDate As Variant  (read-only)`  
  Returns the end of the filtering date range (equal to the StartDate property if range is a single day). Read-only Variant.
- `FilterType As XlPivotFilterType  (read-only)`  
  Returns the type of date filter. Read-only XlPivotFilterType.
- `FilterValue1 As Variant  (read-only)`  
  Returns the first value associated with the date filter (semantics vary by filter type). Read-only Variant.
- `FilterValue2 As Variant  (read-only)`  
  Returns the second value associated with the date filter (semantics vary by filter type). Read-only Variant.
- `SingleRangeFilterState As Boolean  (read-only)`  
  True when the filtering state is a contiguous date range; otherwise, False. Read-only Boolean.

## Methods (1)

- `SetFilterDateRange(StartDate As Variant, EndDate As Variant) As XlFilterStatus`  
  Sets the timeline's filter.
    - `StartDate As Variant` (required): The start of the filtering date range.
    - `EndDate As Variant` (required): The end of the filtering date range.
