# IPicture

**Type:** Interface  
**Library:** OLE Automation  
**GUID:** {7BF80980-BF32-101A-8BBB-00AA00300CAB}  

Picture Object

## Properties (8)

- `Handle As HRESULT  (read-only)`
- `hPal As HRESULT  (read/write)`
- `Type As HRESULT  (read-only)`
- `Width As HRESULT  (read-only)`
- `Height As HRESULT  (read-only)`
- `CurDC As HRESULT  (read-only)`
- `KeepOriginalFormat As HRESULT  (read/write)`
- `Attributes As HRESULT  (read-only)`

## Methods (5)

- `Render(hdc As Long, x As Long, y As Long, cx As Long, cy As Long, xSrc As OLE_XPOS_HIMETRIC, ySrc As OLE_YPOS_HIMETRIC, cxSrc As OLE_XSIZE_HIMETRIC, cySrc As OLE_YSIZE_HIMETRIC, prcWBounds As void)`
- `SelectPicture(hdcIn As Long, phdcOut As Long, phbmpOut As OLE_HANDLE)`
- `PictureChanged()`
- `SaveAsFile(pstm As void, fSaveMemCopy As Boolean, pcbSize As Long)`
- `SetHdc(hdc As OLE_HANDLE)`
