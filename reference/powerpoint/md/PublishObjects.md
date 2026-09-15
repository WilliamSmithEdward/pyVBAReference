# PublishObjects

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {914934CF-5A91-11CF-8700-00AA0060263B}  

A collection of PublishObject objects representing the set of complete or partial loaded presentations that are available for publishing to HTML.

**Remarks:** You can specify the content and attributes of the published presentation by setting various properties of the PublishObject object. For example, the SourceTypeproperty defines the portion of a loaded presentation to be published. The RangeStartproperty and the RangeEndproperty specify the range of slides to publish, and the SpeakerNotesproperty designates whether or not to publish the speaker's notes. You cannot add to the PublishObjects collection.

**Example:**

```vba
With ActivePresentation.PublishObjects(1)

    .FileName = "C:\Test\Mallard.htm"

    .SourceType = ppPublishSlideRange

    .RangeStart = 3

    .RangeEnd = 5

    .Publish

End With
```

## Properties (3)

- `Count As Long  (read-only)`  
  Returns the number of objects in the specified collection. Read-only.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.

## Methods (1)

- `Item(Index As Long) As PublishObject`  
  Returns a single PublishObject object from the specified PublishObjects collection.
    - `Index As Long` (required): The index number of the single PublishObject object in the collection to be returned.
