# Player

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {BA72E552-4FF5-48F4-8215-5505F990966F}  

Allows access to playback controls for the associated shape in the current window.

## Properties (4)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `CurrentPosition As Long  (read/write)`  
  Gets or sets a Long that represents the current position of the player in the media. Read/write.
- `State As PpPlayerState  (read-only)`  
  Returns the current state of the player. Read-only.

## Methods (5)

- `Play()`  
  Begins playback for the specified media.
- `Pause()`  
  Pauses the media.
- `Stop()`  
  Stops the media playing in the specified object.
- `GoToNextBookmark()`  
  Goes to the next bookmark.
- `GoToPreviousBookmark()`  
  Goes to the previous bookmark.
