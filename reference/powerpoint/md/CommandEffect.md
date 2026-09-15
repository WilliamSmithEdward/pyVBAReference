# CommandEffect

**Type:** Dispatch Interface  
**Library:** Microsoft PowerPoint 16.0 Object Library  
**GUID:** {914934EF-5A91-11CF-8700-00AA0060263B}  

Represents a command effect for an animation behavior. You can send events, call functions, and send OLE verbs to embedded objects using this object.

**Remarks:** Use the CommandEffect property of the AnimationBehavior object to return a CommandEffect object. Command effects can be changed using the CommandEffect object's Command and Type properties.

**Example:**

```vba
Set bhvEffect = effectNew.Behaviors.Add(msoAnimTypeCommand)



    With bhvEffect.CommandEffect

         .Type = msoAnimCommandTypeVerb

         .Command = Play

    End With
```

## Properties (5)

- `Application As Application  (read-only)`  
  Returns an Application object that represents the creator of the specified object.
- `Parent As Object  (read-only)`  
  Returns the parent object for the specified object.
- `Type As MsoAnimCommandType  (read/write)`  
  Represents the type of animation. Read/write.
- `Command As String  (read/write)`  
  Represents the command to be executed for the command effect. Read/write.
- `bookmark As String  (read/write)`  
  Sets or returns the bookmark of the specified object. Read/write.
