# EffectParameter

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C03CF-0000-0000-C000-000000000046}  

Describes a single Picture Effect parameter.

**Remarks:** Picture Effects are processed as a chain composed of individual items that are applied in sequence to create the final composited image. An Effects chain will allow an effect to be added to the chain, reordered, or removed from the chain. Effect parameters specify properties of those effects.

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

## Properties (4)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the EffectParameter object. Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the EffectParameter object was created. Read-only.
- `Name As String  (read-only)`  
  Retrieves the string name of the EffectParameter parameter. Read-only.
- `Value As Variant  (read/write)`  
  Retrieves or sets the value of the EffectParameter object. Read/write.
