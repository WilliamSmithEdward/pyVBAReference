# Broadcast

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {BA72E558-4FF5-48F4-8215-5505F990966F}  

A Broadcast Documents library that is used to store documents being presented.

## Properties (8)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object. Read-only.
- `AttendeeUrl As String  (read-only)`  
  Provides the attendee URL that the presenter must share with remote attendees. Read-only.
- `IsBroadcasting As Boolean  (read-only)`  
  Indicates whether the file is in the Broadcast Documents library. Read-only.
- `State As MsoBroadcastState  (read-only)`  
  Returns an MSOBroadcastState constant that describes the current broadcast state. Read-only.
- `Capabilities As Long  (read-only)`  
  Returns a Long that represents the capabilities of the specified broadcast. Read-only.
- `SessionID As String  (read-only)`  
  If a broadcast is in progress, returns a String that specifies the ID of the session. Read-only.
- `PresenterServiceUrl As String  (read-only)`  
  If a broadcast is in progress, returns a String that represents the URL of the presenter service hosting the broadcast. Read-only.

## Methods (5)

- `Start(serverUrl As String)`  
  Returns a list of supported functionality and the maximum allowed file size.
    - `serverUrl As String` (required): The URL of the hosting Web service.
- `End()`  
  Elevates to the system to delete the document from the Broadcast Documents library.
- `Pause()`  
  Pauses the specified broadcast.
- `Resume()`  
  Resumes the specified broadcast.
- `AddMeetingNotes(notesUrl As String, notesWacUrl As String)`  
  Adds shared meeting notes for the specified broadcast that are accessible to attendees who use either Microsoft OneNote 2013 rich client or web app.
