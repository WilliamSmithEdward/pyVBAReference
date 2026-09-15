# Broadcast

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {B67DE22C-BC01-4A73-A99B-070D1B5A795D}  

Represents a Windows Live Broadcast Service broadcast session.

## Properties (7)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified Broadcast object.
- `AttendeeUrl As String  (read-only)`  
  If a broadcast is in progress, returns a String that represents the URL of the attendee link. Read-only.
- `State As MsoBroadcastState  (read-only)`  
  Returns an MSOBroadcastState constant that describes the current broadcast state. Read-only.
- `Capabilities As Long  (read-only)`  
  Returns a Long that represents the capabilities of the specified broadcast. Read-only.
- `PresenterServiceUrl As String  (read-only)`  
  If a broadcast is in progress, returns a String that represents the URL of the presenter service hosting the broadcast. Read-only.
- `SessionID As String  (read-only)`  
  If a broadcast is in progress, returns a String that specifies the ID of the session. Read-only.

## Methods (5)

- `Start(serverUrl As String)`  
  Initiates the specified broadcast session.
    - `serverUrl As String` (required): The URL of the broadcast server.
- `Pause()`  
  Pauses the specified broadcast.
- `Resume()`  
  Resumes the specified broadcast.
- `End()`  
  Ends the specified broadcast session.
- `AddMeetingNotes(notesUrl As String, notesWacUrl As String)`  
  Adds shared meeting notes for the specified broadcast that are accessible to attendees who use either Microsoft OneNote 2013 rich client or web app.
    - `notesUrl As String` (required): Specifies the URL where the shared meeting notes are stored, for attendees using the Microsoft OneNote 2013 rich client.
    - `notesWacUrl As String` (required): Specifies the URL where the shared meeting notes are stored, for attendees using the Microsoft OneNote 2013 web access client.
