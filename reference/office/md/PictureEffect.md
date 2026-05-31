# PictureEffect

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C03D1-0000-0000-C000-000000000046}  

Represents a picture effect.

**Remarks:** Picture effects are processed as a chain composed of individual items that are applied in sequence to create the final composited image. An effects chain will allow an effect to be added to the chain, reordered, or removed from the chain.

**Example:**

```vba
Sub PictureEffectSample()
' Setup a slide with one picture shape.
With ActivePresentation.Slides(1).Shapes(1).Fill.PictureEffects

 ' Insert a 150% Saturation effect.
 .Insert(msoEffectSaturation).EffectParameters(1).Value = 1.5

 ' Insert Brightness/Contrast effect and set values to -50% Brightness and +25% Contrast.
 Dim brightnessContrast As PictureEffect
 Set brightnessContrast = .Insert(msoEffectBrightnessContrast)
 brightnessContrast.EffectParameters(1).Value = -0.5
 brightnessContrast.EffectParameters(2).Value = 0.25

 ' Remove all Picture effects.
 While .Count > 0
 .Delete (1)
 Wend

End With
End Sub
```

## Properties (6)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the PictureEffect object. Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the PictureEffect object was created. Read-only.
- `Type As MsoPictureEffectType  (read-only)`  
  Specifies the type of PictureEffect. Read-only.
- `Position As Long  (read/write)`  
  Specifies the position of a picture effect in a chain of composite effects. Read/write.
- `EffectParameters As EffectParameters  (read-only)`  
  Returns an EffectParameter object. Read-only.
- `Visible As MsoTriState  (read/write)`  
  Gets or sets a Boolean value representing the visible state of the picture effect. Read/write.

## Methods (1)

- `Delete()`  
  Deletes a PictureEffect object.
