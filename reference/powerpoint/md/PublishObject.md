# PublishObject

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {914934D0-5A91-11CF-8700-00AA0060263B}  

Represents a complete or partial loaded presentation that is available for publishing to HTML. The PublishObject object is a member of the PublishObjects collection.

**Remarks:** You can specify the content and attributes of the published presentation by setting various properties of the PublishObject object. For example, the SourceTypeproperty defines the portion of a loaded presentation to be published. The RangeStartproperty and the RangeEndproperty specify the range of slides to publish, and the SpeakerNotesproperty designates whether or not to publish the speaker's notes.

**Example:**

```vba
With Presentations(2).PublishObjects(1)

    .FileName = "C:\Test\Mallard.htm"

    .SourceType = ppPublishSlideRange

    .RangeStart = 3

    .RangeEnd = 5

    .Publish

End With
```

## Properties (9)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `HTMLVersion As PpHTMLVersion  (read/write)`  
  Returns or sets the version of HTML for a published presentation. Read/write.
- `SourceType As PpPublishSourceType  (read/write)`  
  Returns or sets the source type of the presentation to be published to HTML. Read/write.
- `RangeStart As Long  (read/write)`  
  Returns or sets the number of the first slide in a range of slides you are publishing as a Web presentation. Read/write.
- `RangeEnd As Long  (read/write)`  
  Returns or sets the number of the last slide in a range of slides you are publishing as a Web presentation. Read/write.
- `SlideShowName As String  (read/write)`  
  Returns or sets the name of the custom slide show to be published as a Web presentation. Read/write.
- `SpeakerNotes As MsoTriState  (read/write)`  
  Determines whether speaker notes are to be published with the presentation. Read/write.
- `FileName As String  (read/write)`  
  Returns or sets the path and file name of the Web presentation created when all or part of the active presentation is published. Read/write.

## Methods (1)

- `Publish()`  
  Creates a Web presentation (HTML format) from any loaded presentation. You can view the published presentation in a web browser.
