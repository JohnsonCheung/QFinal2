Attribute VB_Name = "MVb_Str_Ens"
Option Explicit
Function EnsSfx$(A, Sfx$)
If HasSfx(A, Sfx) Then EnsSfx = A: Exit Function
EnsSfx = A & Sfx
End Function

Function EnsSfxDot$(A)
EnsSfxDot = EnsSfx(A, ".")
End Function

Function EnsSfxSC$(A)
EnsSfxSC = EnsSfx(A, ";")
End Function
