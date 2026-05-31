# Speech

**Type:** Dispatch Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00024466-0000-0000-C000-000000000046}  

Contains methods and properties that pertain to speech.

**Remarks:** Use the Speech property of the Application object to return a Speech object.

**Example:**

```vba
Sub UseSpeech()

 Application.Speech.Speak "Hello"

End Sub()
```

## Properties (2)

- `Direction As XlSpeakDirection  (read/write)`  
  Returns or sets the order in which the cells will be spoken. The value of the Direction property is an XlSpeakDirection constant. Read/write.
- `SpeakCellOnEnter As Boolean  (read/write)`  
  Microsoft Excel supports a mode where the active cell is spoken when the Enter key is pressed or when the active cell is finished being edited. Setting the SpeakCellOnEnter property to True turns this mode on. False turns this mode off. Read/write Boolean.

## Methods (1)

- `Speak(Text As String, [SpeakAsync As Variant], [SpeakXML As Variant], [Purge As Variant])`  
  Microsoft Excel plays back the text string that is passed as an argument.
    - `Text As String` (required): The text to be spoken.
    - `SpeakAsync As Variant` (optional): True causes the text to be spoken asynchronously (the method will not wait for the text to be spoken). False causes the text to be spoken synchronously (the method waits for the text to be spoken before continuing). The default is False.
    - `SpeakXML As Variant` (optional): True causes the text to be interpreted as XML. False causes the text to not be interpreted as XML, so any XML tags are read and not interpreted. The default is False.
    - `Purge As Variant` (optional): True causes current speech to be terminated and any buffered text to be purged before text is spoken. False does not cause the current speech to be terminated and does not purge the buffered text before text is spoken. The default is False.
