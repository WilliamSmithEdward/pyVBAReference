# EffectParameters

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C03D0-0000-0000-C000-000000000046}  

Represents a collection of EffectParameter objects.

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

## Properties (5)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the EffectParameters object. Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the EffectParameters object was created. Read-only.
- `Item As EffectParameter  (read-only)`  
  Retrieves an EffectParameter object at the specified index or with the specified unique Id. Read-only.
- `Count As Long  (read-only)`  
  Retrieves the count of the number of EffectParameter objects contained within the EffectParameters collection. Read-only.
- `_NewEnum As IUnknown  (read-only)`
