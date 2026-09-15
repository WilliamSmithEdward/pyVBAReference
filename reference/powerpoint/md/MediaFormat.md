# MediaFormat

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {BA72E550-4FF5-48F4-8215-5505F990966F}  

Contains methods and properties that allow access to and control over audio and video media.

## Properties (19)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `Volume As Single  (read/write)`  
  Gets or sets the volume of the media. Read/write.
- `Muted As Boolean  (read/write)`  
  Returns whether audio playback of the media is muted. Read/write
- `Length As Long  (read-only)`  
  Returns the total length of the media in milliseconds. Read-only.
- `StartPoint As Long  (read/write)`  
  Gets or sets the start point of the trim region. Read/write.
- `EndPoint As Long  (read/write)`  
  Gets or sets the time of the end point of the trim region of the media. Read/write.
- `FadeInDuration As Long  (read/write)`  
  Gets or sets the duration of the fade in of the media, in milliseconds. Read/write.
- `FadeOutDuration As Long  (read/write)`  
  Gets or sets the duration of the fadeout of the media, in milliseconds. Read/write.
- `MediaBookmarks As MediaBookmarks  (read-only)`  
  Returns a MediaBookmarks collection that represents the media bookmarks associated with the specified object. Read-only.
- `ResamplingStatus As PpMediaTaskStatus  (read-only)`  
  Returns the resampling task status. Read-only.
- `IsLinked As Boolean  (read-only)`  
  Returns whether the media file is linked. Read-only.
- `IsEmbedded As Boolean  (read-only)`  
  Returns whether the media file is embedded. Read-only.
- `AudioSamplingRate As Long  (read-only)`  
  Returns the audio sampling rate per second for the media. Read-only.
- `VideoFrameRate As Long  (read-only)`  
  Returns the video frame rate per second of the media. Read-only.
- `SampleHeight As Long  (read-only)`  
  Returns the resolution height of the sample media. Read-only.
- `SampleWidth As Long  (read-only)`  
  Returns the resolution width of the media sample. Read-only.
- `VideoCompressionType As String  (read-only)`  
  Returns a String that represents the video compression format of the media. Read-only.
- `AudioCompressionType As String  (read-only)`  
  Returns the audio compression format. Read-only.

## Methods (4)

- `SetDisplayPicture(Position As Long)`  
  Sets the display picture at the specified time position.
    - `Position As Long` (required): The time position in the media to set the display picture.
- `SetDisplayPictureFromFile(FilePath As String)`  
  Sets the display picture from a picture file.
    - `FilePath As String` (required): The path to the display picture file.
- `Resample([Trim As Boolean], [SampleHeight As Long], [SampleWidth As Long], [VideoFrameRate As Long], [AudioSamplingRate As Long], [VideoBitRate As Long])`  
  Adds the current media object to the queue and begins resampling, based on the specified parameters.
    - `Trim As Boolean` (optional): Whether to trim the sample.
    - `SampleHeight As Long` (optional): The sample resolution height.
    - `SampleWidth As Long` (optional): The sample resolution width.
    - `VideoFrameRate As Long` (optional): The video frame rate, in frames per second.
    - `AudioSamplingRate As Long` (optional): The audio sampling rate, in bits per second.
    - `VideoBitRate As Long` (optional): The video bit rate, in bits per second.
- `ResampleFromProfile([profile As PpResampleMediaProfile])`  
  Adds the current media object to the queue and begins resampling base on the specified profile.
    - `profile As PpResampleMediaProfile` (optional): The resample media profile to use.
