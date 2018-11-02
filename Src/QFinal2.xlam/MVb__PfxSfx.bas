Attribute VB_Name = "MVb__PfxSfx"
Option Explicit
Option Compare Text
Function AddPfx$(A$, Pfx$)
AddPfx = Pfx & A
End Function

Function AddPfxSfx$(A$, Pfx$, Sfx$)
AddPfxSfx = Pfx & A & Sfx
End Function

Function AddSfx$(A$, Sfx$)
AddSfx = A & Sfx
End Function

Function HasPfxAy(A, PfxAy0, Optional Cmp As VbCompareMethod = vbTextCompare) As Boolean
Dim I
For Each I In CvNy(PfxAy0)
   If HasPfx(A, I, Cmp) Then HasPfxAy = True: Exit Function
Next
End Function

Function HasPfxAyCasSen(A, PfxAy0) As Boolean
HasPfxAyCasSen = HasPfxAy(A, PfxAy0, vbTextCompare)
End Function

Function HasPfx(A, Pfx, Optional Cmp As VbCompareMethod = vbTextCompare) As Boolean
HasPfx = StrComp(Left(A, Len(Pfx)), Pfx, Cmp) = 0
End Function

Function HasPfxAp(A, ParamArray PfxAp()) As Boolean
Dim Av(): Av = PfxAp
HasPfxAp = HasPfxAy(A, Av)
End Function

Function HasPfxSpc(A, Pfx) As Boolean
HasPfxSpc = HasPfx(A, Pfx & " ")
End Function
Function AyAddPfx(A, Pfx) As String()
If Sz(A) = 0 Then Exit Function
Dim O$(), U&, J&
U = UB(A)
ReDim O(U)
For J = 0 To U
    O(J) = Pfx & A(J)
Next
AyAddPfx = O
End Function

Function AyAddPfxSfx(A, Pfx, Sfx) As String()
Dim O$(), J&, U&
If Sz(A) = 0 Then Exit Function
U = UB(A)
ReDim Preserve O(U)
For J = 0 To U
    O(J) = Pfx & A(J) & Sfx
Next
AyAddPfxSfx = O
End Function

Function AyAddSfx(A, Sfx) As String()
If Sz(A) = 0 Then Exit Function
Dim O$(), J&, U&
U = UB(A)
ReDim Preserve O(U)
For J = 0 To U
    O(J) = A(J) & Sfx
Next
AyAddSfx = O
End Function

Function AyIsAllEleHasPfx(A, Pfx$) As Boolean
If Sz(A) = 0 Then Exit Function
Dim I
For Each I In A
   If Not HasPfx(I, Pfx) Then Exit Function
Next
AyIsAllEleHasPfx = True
End Function


Function AyAddCommaSpcSfxExlLas(A) As String()
Dim X, J, U%
U = UB(A)
For Each X In AyNz(A)
    If J = U Then
        Push AyAddCommaSpcSfxExlLas, X
    Else
        Push AyAddCommaSpcSfxExlLas, X & ", "
    End If
    J = J + 1
Next
End Function

Function HasSfx(A, Sfx, Optional Cmp As VbCompareMethod = vbTextCompare) As Boolean
HasSfx = StrComp(Right(A, Len(Sfx)), Sfx, Cmp) = 0
End Function


Function SyIsAllEleHasPfx(A$(), Pfx$) As Boolean
If Sz(A) = 0 Then Exit Function
Dim I
For Each I In A
   If Not HasPfx(CStr(I), Pfx) Then Exit Function
Next
SyIsAllEleHasPfx = True
End Function

Function AyAddIxPfx(A) As String()
Dim I, J&
For Each I In AyNz(A)
    PushI AyAddIxPfx, J & ": " & I
    J = J + 1
Next
End Function
