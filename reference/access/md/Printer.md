# Printer

**Type:** Class  
**Library:** Microsoft Access 16.0 Object Library  
**GUID:** {DBC5175E-A8ED-11D3-A0DD-00C04F68712B}  

A Printer object corresponds to a printer available on your system.

**Remarks:** A Printer object is a member of the Printers collection. To return a reference to a particular Printer object in the Printers collection, use any of the following syntax forms. Use the properties of the Printer object to set the printing characteristics for any of the printers available on your system. Use the ColorMode, Copies, Duplex, Orientation, PaperBin, PaperSize, and PrintQuality properties to specify print settings for a particular printer. Use the LeftMargin, RightMargin, TopMargin, BottomMargin, ColumnSpacing, RowSpacing, DataOnly, DefaultSize, ItemLayout, ItemsAcross, ItemSizeHeight, and ItemSizeWidth properties to specify how Microsoft Access should format the appearance of data on printed pages. Use the DeviceName, DriverName, and Port properties to return system information about a particular printer.

**Example:**

```vba
Dim prtFirst As Printer

Set prtFirst = Application.Printers(0)

With prtFirst
 MsgBox "Device name: " & .DeviceName & vbCr _
 & "Driver name: " & .DriverName & vbCr _
 & "Port: " & .Port
End With
```

## Properties (22)

- `ColorMode As AcPrintColor  (read/write)`  
  Returns or sets an AcPrintColor constant representing whether the specified printer should print output in color or monochrome. Read/write.
- `Copies As Long  (read/write)`  
  Returns or sets a Long indicating the number of copies to be printed. Read/write.
- `DeviceName As String  (read-only)`  
  Returns a String indicating the name of the specified printer device. Read-only.
- `DriverName As String  (read-only)`  
  Returns a String indicating the name of the driver used by the specified printer. Read-only.
- `Duplex As AcPrintDuplex  (read/write)`  
  Returns or sets an AcPrintDuplex constant indicating how the specified printer handles duplex printing. Read/write.
- `Orientation As AcPrintOrientation  (read/write)`  
  Gets or sets an AcPrintOrientation constant that represents the print orientation. Read/write.
- `PaperBin As AcPrintPaperBin  (read/write)`  
  Returns or sets an AcPrintPaperBin constant indicating which paper bin the specified printer should use. Read/write.
- `PaperSize As AcPrintPaperSize  (read/write)`  
  Returns or sets an AcPrintPaperSize constant indicating the paper size to use when printing. Read/write.
- `Port As String  (read-only)`  
  Returns a String indicating the port name of the specified printer. Read-only.
- `PrintQuality As AcPrintObjQuality  (read/write)`  
  Returns or sets an AcPrintObjQuality constant indicating the resolution at which the specified printer should print jobs. Read/write.
- `LeftMargin As Long  (read/write)`  
  Along with the TopMargin, RightMargin, and BottomMargin properties, specifies the margins for a printed page. Read/write Long.
- `RightMargin As Long  (read/write)`  
  Along with the TopMargin, LeftMargin, and BottomMargin properties, specifies the margins for a printed page. Read/write Long.
- `TopMargin As Long  (read/write)`  
  Along with the LeftMargin, RightMargin, and BottomMargin properties, specifies the margins for a printed page. Read/write Long.
- `BottomMargin As Long  (read/write)`  
  Along with the TopMargin, RightMargin, and LeftMargin properties, specifies the margins for a printed page. Read/write Long.
- `DataOnly As Boolean  (read/write)`  
  True if Microsoft Access prints only the data from a table or query in Datasheet view and not the labels, control borders, gridlines, and display graphics. Read/write Boolean.
- `ItemsAcross As Long  (read/write)`  
  Returns or sets a Long indicating the number of columns to print across a page for multiple-column reports or labels. Read/write.
- `RowSpacing As Long  (read/write)`  
  Returns or sets a Long indicating the horizontal space between detail sections in twips. Read/write.
- `ColumnSpacing As Long  (read/write)`  
  Returns or sets a Long representing the vertical space between detail sections in twips. Read/write.
- `DefaultSize As Boolean  (read/write)`  
  True if the size of the detail section in Design view is used for printing; otherwise, the values of the ItemSizeHeight and ItemSizeWidth properties are used. Read/write Boolean.
- `ItemSizeWidth As Long  (read/write)`  
  Returns or sets a Long indicating the width of the detail section of a form or report in twips. Read/write.
- `ItemSizeHeight As Long  (read/write)`  
  Returns or sets a Long indicating the height of the detail section of a form or report in twips. Read/write.
- `ItemLayout As AcPrintItemLayout  (read/write)`  
  Returns or sets an AcPrintItemLayout constant indicating whether the printer lays columns across, then down, or down, then across. Read/write.
