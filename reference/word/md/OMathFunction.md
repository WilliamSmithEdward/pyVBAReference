# OMathFunction

**Type:** Dispatch Interface  
**Library:** Microsoft Word 16.0 Object Library  
**GUID:** {F1F37152-1DB1-4901-AD9A-C740F99464B4}  

Represents a mathematical function or structure that Microsoft Word supports, such as fractions, integrals, sums, and radicals. The OMathFunction object is a member of the OMathFunctions collection.

## Properties (26)

- `Type As WdOMathFunctionType  (read-only)`  
  Returns a WdOMathFunctionType constant that represents the type of function. Read-only.
- `Application As Application  (read-only)`  
  Returns an Application object that represents the Microsoft Word application.
- `Creator As Long  (read-only)`  
  Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only Long.
- `Parent As Object  (read-only)`  
  Returns an Object that represents the parent object of the specified OMathFunction object.
- `Range As Range  (read-only)`  
  Returns a Range object that represents the portion of a document that is contained in the specified object. Read-only.
- `Args As OMathArgs  (read-only)`  
  Returns an OMathArgs object that represents the arguments for an equation. Read-only.
- `Acc As OMathAcc  (read-only)`  
  Returns an OMathAcc object that represents a base character with a combining accent mark. Read-only.
- `Bar As OMathBar  (read-only)`  
  Returns an OMathBar object that represents the mathematical overbar for an object. Read-only.
- `Box As OMathBox  (read-only)`  
  Returns an OMathBox object that represents an invisible box around an equation or part of an equation to which you can apply properties that affect the mathematical or formatting properties, such as line breaks. Read-only.
- `BorderBox As OMathBorderBox  (read-only)`  
  Returns an OMathBorderBox object that represents a border drawn around an equation or part of an equation. The BorderBox object can also be used to draw horizontal, vertical, and diagonal strikethrough lines through the BorerBox arguments. Read-only.
- `Delim As OMathDelim  (read-only)`  
  Returns an OMathDelim object that represents the delimiter function. Read-only.
- `EqArray As OMathEqArray  (read-only)`  
  Returns an OMathEqArray object that represents an equation array function. Read-only.
- `Frac As OMathFrac  (read-only)`  
  Returns an OMathFrac object that represents a fraction. Read-only.
- `Func As OMathFunc  (read-only)`  
  Returns an OMathFunc object that represents a type of mathematical function that consists of a function name, such as sin or cos, and an argument. Read-only.
- `GroupChar As OMathGroupChar  (read-only)`  
  Returns an OMathGroupChar object that represents a horizontal character placed above or below text in an equation, often with the purpose of grouping the text visually. Read-only.
- `LimLow As OMathLimLow  (read-only)`  
  Returns an OMathLimLow object that represents the lower limit for a function. Read-only.
- `LimUpp As OMathLimUpp  (read-only)`  
  Returns an OMathLimUpp object that represents upper limit function. Read-only.
- `Mat As OMathMat  (read-only)`  
  Returns an OMathMat object that represents a mathematical matrix. Read-only.
- `Nary As OMathNary  (read-only)`  
  Returns an OMathNary object that represents the n-ary operation. Read-only.
- `Phantom As OMathPhantom  (read-only)`  
  Returns an OMathPhantom object that represents an object used for advanced layout of an equation. Read-only.
- `ScrPre As OMathScrPre  (read-only)`  
  Returns an OMathScrPre object that represents a superscript and subscript to the left of the base. Read-only.
- `Rad As OMathRad  (read-only)`  
  Returns an OMathRad object that represents the mathematical radical function. Read-only.
- `ScrSub As OMathScrSub  (read-only)`  
  Represents an OMathScrSub object that represents the mathematical subscript function. Read-only.
- `ScrSubSup As OMathScrSubSup  (read-only)`  
  Returns an OMathScrSubSup object that represents a mathematical subscript-superscript object that consists of a base, a subscript, and a superscript. Read-only.
- `ScrSup As OMathScrSup  (read-only)`  
  Returns an OMathScrSup object that represents the mathematical superscript function. Read-only.
- `OMath As OMath  (read-only)`  
  Returns an OMath object that represents the equation. Read-only.

## Methods (1)

- `Remove() As OMathFunction`  
  Removes a function from the collection of functions contained within an equation and returns an OMathFunction object.
