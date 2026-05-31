# IScenario

**Type:** Interface  
**Library:** Microsoft Excel 16.0 Object Library  
**GUID:** {00020897-0001-0000-C000-000000000046}  

## Properties (10)

- `Application As HRESULT  (read-only)`
- `Creator As HRESULT  (read-only)`
- `Parent As HRESULT  (read-only)`
- `ChangingCells As HRESULT  (read-only)`
- `Comment As HRESULT  (read/write)`
- `Hidden As HRESULT  (read/write)`
- `Index As HRESULT  (read-only)`
- `Locked As HRESULT  (read/write)`
- `Name As HRESULT  (read/write)`
- `Values As HRESULT  (read-only)`

## Methods (3)

- `ChangeScenario(ChangingCells As Variant, [Values As Variant], RHS As Variant)`
- `Delete(RHS As Variant)`
- `Show(RHS As Variant)`
