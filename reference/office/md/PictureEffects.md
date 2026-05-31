# PictureEffects

**Type:** Dispatch Interface  
**Library:** Microsoft Office 16.0 Object Library  
**GUID:** {000C03D2-0000-0000-C000-000000000046}  

Represents a collection of PictureEffect objects.

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

## Properties (5)

- `Application As Object  (read-only)`  
  Gets an Application object that represents the container application for the PictureEffects object. Read-only.
- `Creator As Long  (read-only)`  
  Gets a 32-bit integer that indicates the application in which the PictureEffects object was created. Read-only.
- `Item As PictureEffect  (read-only)`  
  Retrieves a PictureEffect object at the specified index. Read-only.
- `Count As Long  (read-only)`  
  Retrieves the count of the number of PictureEffect objects contained within the PictureEffects collection. Read-only.
- `_NewEnum As IUnknown  (read-only)`

## Methods (2)

- `Insert(EffectType As MsoPictureEffectType, [Position As Long]) As PictureEffect`  
  Inserts a picture effect in a chain of composite effects.
    - `EffectType As MsoPictureEffectType` (required): An enumeration specifying the type of picture effect.
    - `Position As Long` (optional): The position of the effect in the composite chain of picture effects.
- `Delete([Index As Long])`  
  Deletes a PictureEffect object from the collection.
    - `Index As Long` (optional): The index number of the PictureEffect object to delete.
