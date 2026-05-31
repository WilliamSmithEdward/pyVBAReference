# PageSetup

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {000208B4-0000-0000-C000-000000000046}  

Represents the page setup description.

**Remarks:** The PageSetup object contains all page setup attributes (left margin, bottom margin, paper size, and so on) as properties.

**Example:**

```vba
With Worksheets("Sheet1")
 .PageSetup.Orientation = xlLandscape
 .PrintOut
End With
```

## Properties (48)

- `Application As Application  (read-only)`  
  When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
- `Creator As XlCreator  (read-only)`  
  Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `BlackAndWhite As Boolean  (read/write)`  
  True if elements of the document will be printed in black and white. Read/write Boolean.
- `BottomMargin As Double  (read/write)`  
  Returns or sets the size of the bottom margin, in points. Read/write Double.
- `CenterFooter As String  (read/write)`  
  Center aligns the footer information in the PageSetup object. Read/write String.
- `CenterHeader As String  (read/write)`  
  Center aligns the header information in the PageSetup object. Read/write String.
- `CenterHorizontally As Boolean  (read/write)`  
  True if the sheet is centered horizontally on the page when it's printed. Read/write Boolean.
- `CenterVertically As Boolean  (read/write)`  
  True if the sheet is centered vertically on the page when it's printed. Read/write Boolean.
- `Draft As Boolean  (read/write)`  
  True if the sheet will be printed without graphics. Read/write Boolean.
- `FirstPageNumber As Long  (read/write)`  
  Returns or sets the first page number that will be used when this sheet is printed. If xlAutomatic, Microsoft Excel chooses the first page number. The default is xlAutomatic (Constants). Read/write Long.
- `FitToPagesTall As Variant  (read/write)`  
  Returns or sets the number of pages tall that the worksheet will be scaled to when it's printed. Applies only to worksheets. Read/write Variant.
- `FitToPagesWide As Variant  (read/write)`  
  Returns or sets the number of pages wide that the worksheet will be scaled to when it's printed. Applies only to worksheets. Read/write Variant.
- `FooterMargin As Double  (read/write)`  
  Returns or sets the distance from the bottom of the page to the footer, in points. Read/write Double.
- `HeaderMargin As Double  (read/write)`  
  Returns or sets the distance from the top of the page to the header, in points. Read/write Double.
- `LeftFooter As String  (read/write)`  
  Returns or sets the alignment of text on the left footer of a workbook or section.
- `LeftHeader As String  (read/write)`  
  Returns or sets the alignment of text on the left header of a workbook or section.
- `LeftMargin As Double  (read/write)`  
  Returns or sets the size of the left margin, in points. Read/write Double.
- `Order As XlOrder  (read/write)`  
  Returns or sets an XlOrder value that represents the order that Microsoft Excel uses to number pages when printing a large worksheet.
- `Orientation As XlPageOrientation  (read/write)`  
  Returns or sets an XlPageOrientation value that represents the portrait or landscape printing mode.
- `PaperSize As XlPaperSize  (read/write)`  
  Returns or sets the size of the paper. Read/write XlPaperSize.
- `PrintArea As String  (read/write)`  
  Returns or sets the range to be printed as a String using A1-style references in the language of the macro. Read/write String.
- `PrintGridlines As Boolean  (read/write)`  
  True if cell gridlines are printed on the page. Applies only to worksheets. Read/write Boolean.
- `PrintHeadings As Boolean  (read/write)`  
  True if row and column headings are printed with this page. Applies only to worksheets. Read/write Boolean.
- `PrintNotes As Boolean  (read/write)`  
  True if cell notes are printed as end notes with the sheet. Applies only to worksheets. Read/write Boolean.
- `PrintQuality As Variant  (read/write)`  
  Returns or sets the print quality. Read/write Variant.
- `PrintTitleColumns As String  (read/write)`  
  Returns or sets the columns that contain the cells to be repeated on the left side of each page, as a String in A1-style notation in the language of the macro. Read/write String.
- `PrintTitleRows As String  (read/write)`  
  Returns or sets the rows that contain the cells to be repeated at the top of each page, as a String in A1-style notation in the language of the macro. Read/write String.
- `RightFooter As String  (read/write)`  
  Returns or sets the distance (in points) between the right edge of the page and the right boundary of the footer. Read/write String.
- `RightHeader As String  (read/write)`  
  Returns or sets the right part of the header. Read/write String.
- `RightMargin As Double  (read/write)`  
  Returns or sets the size of the right margin, in points. Read/write Double.
- `TopMargin As Double  (read/write)`  
  Returns or sets the size of the top margin, in points. Read/write Double.
- `Zoom As Variant  (read/write)`  
  Returns or sets a Variant value that represents a percentage (between 10 and 400 percent) by which Microsoft Excel will scale the worksheet for printing.
- `PrintComments As XlPrintLocation  (read/write)`  
  Returns or sets the way comments are printed with the sheet. Read/write XlPrintLocation.
- `PrintErrors As XlPrintErrors  (read/write)`  
  Sets or returns an XlPrintErrors constant specifying the type of print error displayed. This feature allows users to suppress the display of error values when printing a worksheet. Read/write.
- `CenterHeaderPicture As Graphic  (read-only)`  
  Returns a Graphic object that represents the picture for the center section of the header. Used to set attributes about the picture.
- `CenterFooterPicture As Graphic  (read-only)`  
  Returns a Graphic object that represents the picture for the center section of the footer. Used to set attributes about the picture.
- `LeftHeaderPicture As Graphic  (read-only)`  
  Returns a Graphic object that represents the picture for the left section of the header. Used to set attributes about the picture.
- `LeftFooterPicture As Graphic  (read-only)`  
  Returns a Graphic object that represents the picture for the left section of the footer. Used to set attributes about the picture.
- `RightHeaderPicture As Graphic  (read-only)`  
  Returns a Graphic object that represents the picture for the right section of the header. Used to set attributes about the picture.
- `RightFooterPicture As Graphic  (read-only)`  
  Returns a Graphic object that represents the picture for the right section of the footer. Used to set attributes of the picture.
- `OddAndEvenPagesHeaderFooter As Boolean  (read/write)`  
  True if the specified PageSetup object has different headers and footers for odd-numbered and even-numbered pages. Read/write Boolean.
- `DifferentFirstPageHeaderFooter As Boolean  (read/write)`  
  True if a different header or footer is used on the first page. Read/write Boolean.
- `ScaleWithDocHeaderFooter As Boolean  (read/write)`  
  Returns or sets if the header and footer should be scaled with the document when the size of the document changes. Read/write Boolean.
- `AlignMarginsHeaderFooter As Boolean  (read/write)`  
  Returns True for Excel to align the header and the footer with the margins set in the page setup options. Read/write Boolean.
- `Pages As Pages  (read-only)`  
  Returns or sets the count or item number of the pages in the Pages collection.
- `EvenPage As Page  (read-only)`  
  Returns or sets the alignment of text on the even page of a workbook or section.
- `FirstPage As Page  (read-only)`  
  Returns or sets the alignment of text on the first page of a workbook or section.
