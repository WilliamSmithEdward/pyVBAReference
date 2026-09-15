# InlineShapes

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {000209A9-0000-0000-C000-000000000046}  

A collection of InlineShape objects that represent all the inline shapes in a document, range, or selection.

**Remarks:** Use the InlineShapes property to return the InlineShapes collection. The following example converts each inline shape in the active document to a Shape object. Use the New method to create a new picture as an inline shape. Use the AddPicture and AddOLEObject methods to add pictures or OLE objects and link them to a source file. Use the AddOLEControl method to add an ActiveX control. Shape objects are anchored to a range of text but are free-floating and can be positioned anywhere on the page. Use the ConvertToInlineShape method and the ConvertToShape method to convert shapes from one type to the other. You can convert only pictures, OLE objects, and ActiveX controls to inline shapes. The Count property for this collection in a document returns the number of items in the main story only. To count items in other stories use the collection with the Range object. When you open a document created in an earlier version of Word, pictures are converted to inline shapes.

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified InlineShapes object.
- `Count As Long  (read-only)`  
  Returns a Long that represents the number of inline shapes in the collection. Read-only.
- `_NewEnum As IUnknown  (read-only)`

## Methods (12)

- `Item(Index As Long) As InlineShape`  
  Returns an individual InlineShape object in a collection.
    - `Index As Long` (required): The individual object to be returned. Can be a Long indicating the ordinal position of the individual object.
- `AddPicture(FileName As String, [LinkToFile As Variant], [SaveWithDocument As Variant], [Range As Variant]) As InlineShape`  
  Adds a picture to a document. Returns an InlineShape object that represents the picture.
    - `FileName As String` (required): The path and file name of the picture.
    - `LinkToFile As Variant` (optional): True to link the picture to the file from which it was created. False to make the picture an independent copy of the file. The default value is False.
    - `SaveWithDocument As Variant` (optional): True to save the linked picture with the document. The default value is False.
    - `Range As Variant` (optional): The location where the picture will be placed in the text. If the range isn't collapsed, the picture replaces the range; otherwise, the picture is inserted. If this argument is omitted, the picture is placed automatically.
- `AddOLEObject([ClassType As Variant], [FileName As Variant], [LinkToFile As Variant], [DisplayAsIcon As Variant], [IconFileName As Variant], [IconIndex As Variant], [IconLabel As Variant], [Range As Variant]) As InlineShape`  
  Creates an OLE object. Returns the InlineShape object that represents the new OLE object.
    - `ClassType As Variant` (optional): The name of the application used to activate the specified OLE object.
    - `FileName As Variant` (optional): The file from which the object is to be created. If this argument is omitted, the current folder is used. You must specify either the ClassType or FileName argument for the object, but not both.
    - `LinkToFile As Variant` (optional): True to link the OLE object to the file from which it was created. False to make the OLE object an independent copy of the file. If you specified a value for ClassType, the LinkToFile argument must be False. The default value is False.
    - `DisplayAsIcon As Variant` (optional): True to display the OLE object as an icon. The default value is False.
    - `IconFileName As Variant` (optional): The file that contains the icon to be displayed.
    - `IconIndex As Variant` (optional): The index number of the icon within IconFileName. The order of icons in the specified file corresponds to the order in which the icons appear in the Change Icon dialog box (Insert menu, Object dialog box) when the Display as icon check box is selected. The first icon in the file has the index number 0 (zero). If an icon with the given index number doesn't exist in IconFileName, the icon with the index number 1 (the second icon in the file) is used. The default value is 0 (zero).
    - `IconLabel As Variant` (optional): A label (caption) to be displayed beneath the icon.
    - `Range As Variant` (optional): The range where the OLE object will be placed in the text. The OLE object replaces the range, unless the range is collapsed. If this argument is omitted, the object is placed automatically.
- `AddOLEControl([ClassType As Variant], [Range As Variant]) As InlineShape`  
  Creates an ActiveX control (formerly known as an OLE control). Returns the InlineShape object that represents the new ActiveX control.
    - `ClassType As Variant` (optional): The programmatic identifier for the ActiveX control to be created.
    - `Range As Variant` (optional): The range where the ActiveX control will be placed in the text. The ActiveX control replaces the range, if the range isn't collapsed. If this argument is omitted, the Active X control is placed automatically.
- `New(Range As Range) As InlineShape`  
  Inserts an empty, 1-inch-square Word picture object surrounded by a border. This method returns the new graphic as an InlineShape object.
    - `Range As Range` (required): The location of the new graphic.
- `AddHorizontalLine(FileName As String, [Range As Variant]) As InlineShape`  
  Adds a horizontal line based on an image file to the current document.
    - `FileName As String` (required): The file name of the image you want to use for the horizontal line.
    - `Range As Variant` (optional): The range above which Microsoft Word places the horizontal line. If this argument is omitted, Word places the horizontal line above the current selection.
- `AddHorizontalLineStandard([Range As Variant]) As InlineShape`  
  Adds a horizontal line to the current document.
    - `Range As Variant` (optional): The range above which Microsoft Word places the horizontal line. If this argument is omitted, Word places the horizontal line above the current selection.
- `AddPictureBullet(FileName As String, [Range As Variant]) As InlineShape`  
  Adds a picture bullet based on an image file to the current document. Returns an InlineShape object.
    - `FileName As String` (required): The file name of the image you want to use for the picture bullet.
    - `Range As Variant` (optional): The range to which Microsoft Word adds the picture bullet. Word adds the picture bullet to each paragraph in the range. If this argument is omitted, Word adds the picture bullet to each paragraph in the current selection.
- `AddSmartArt(Layout As SmartArtLayout, [Range As Variant]) As InlineShape`  
  Inserts a SmartArt graphic as an inline shape into the active document.
    - `Layout As SmartArtLayout` (required): A SmartArtLayoutobject that specifies the layout for the SmartArt graphic.
    - `Range As Variant` (optional): Specifies the text to which the SmartArt graphic is bound. If Range is specified, the SmartArt graphic is positioned at the beginning of the first paragraph in the range. If this argument is omitted, the range is selected automatically, and the SmartArt graphic is positioned relative to the top and left edges of the page.
- `AddWebVideo(EmbedCode As String, VideoWidth As Variant, VideoHeight As Variant, [PosterFrameImage As Variant], [Url As Variant], [Range As Variant]) As InlineShape`  
  Adds a new web video to the document.
    - `EmbedCode As String` (required): The embed code for the video.
    - `VideoWidth As Variant` (required): An integer that represents the width of the web video in pixels.
    - `VideoHeight As Variant` (required): An integer that represents the height of the web video in pixels.
    - `PosterFrameImage As Variant` (optional): A string that points to the file to use as the poster frame for the web video.
    - `Url As Variant` (optional): The URL to the video.
    - `Range As Variant` (optional): The range at which to insert the web video. If _Range_ is omitted, the current selection is used.
- `AddChart2([Style As Long], [Type As XlChartType], [Range As Variant], [NewLayout As Variant]) As InlineShape`  
  Adds a chart to the document. Returns an InlineShape object that represents the chart and adds it to the specified collection.
    - `Style As Long` (optional): The chart style. Use "-1" to get the default style for the chart type specified in Type.
    - `Range As Variant` (optional): The range where the chart will be placed in the text. The chart replaces the range, unless the range is collapsed. If this argument is omitted, the chart is placed automatically.
    - `NewLayout As Variant` (optional): If _NewLayout_ is true, the chart is inserted by using the new dynamic formatting rules (Title is on, and Legend is on only if there are multiple series).
- `Add3DModel(FileName As String, [LinkToFile As Variant], [SaveWithDocument As Variant], [Range As Variant]) As InlineShape`
