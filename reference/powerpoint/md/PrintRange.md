# PrintRange

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {9149345F-5A91-11CF-8700-00AA0060263B}  

Represents a single range of consecutive slides or pages to be printed.

**Remarks:** The PrintRange object is a member of the PrintRanges collection. The PrintRanges collection contains all the print ranges that have been defined for the specified presentation. You can set print ranges in the PrintRanges collection independent of the RangeType setting; these ranges are retained as long as the presentation they're contained in is loaded. The ranges in the PrintRanges collection are applied when the RangeType property is set to ppPrintSlideRange.

**Example:**

```vba
With ActivePresentation.PrintOptions.Ranges
    If .Count > 0 Then
        With .Item(1)
            MsgBox "Print range 1 starts on slide " & .Start & _
                " and ends on slide " & .End
        End With
    End If
End With
```

## Properties (4)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `Start As Long  (read-only)`  
  Returns the number of the first slide in the range of slides to be printed. Read-only.
- `End As Long  (read-only)`  
  Returns the number of the last slide in the specified print range. Read-only.

## Methods (1)

- `Delete()`  
  Deletes the specified PrintRange object.
