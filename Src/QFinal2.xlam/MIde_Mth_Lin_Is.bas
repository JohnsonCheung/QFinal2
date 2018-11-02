Attribute VB_Name = "MIde_Mth_Lin_Is"
Option Explicit
Function LinIsPrp(A) As Boolean
LinIsPrp = LinMthKd(A) = "Property"
End Function
Private Sub Z_LinIsMth()
GoTo ZZ
Dim A$
A = "Function LinIsMth(A, Optional B As WhMth) As Boolean"
Ept = True
GoSub Tst
Exit Sub
Tst:
    Act = LinIsMth(A)
    C
    Return
ZZ:
Dim L, O$()
For Each L In CurSrc
    If LinIsMth(CStr(L)) Then
        PushI O, L
    End If
Next
Brw O
End Sub


Function IsCdLin(A) As Boolean
Dim L$: L = Trim(A)
If A = "" Then Exit Function
If Left(A, 1) = "'" Then Exit Function
IsCdLin = True
End Function


Function LinIsMth(A) As Boolean
LinIsMth = LinMthKd(A) <> ""
End Function

Function LinIsMthWh(A$, B As WhMth) As Boolean
LinIsMthWh = MthNmBrkIsSel(LinMthNmBrk(A), B)
End Function

Function LinIsPubMth(A) As Boolean
Dim L$: L = A
If Not AyHas(Array("", "Public"), ShfMdy(L)) Then Exit Function
LinIsPubMth = ShfMthTy(L) <> ""
End Function

Function LinIsTstSub(L$) As Boolean
LinIsTstSub = True
If HasPfx(L, "Sub Tst()") Then Exit Function
If HasPfx(L, "Sub Tst()") Then Exit Function
If HasPfx(L, "Friend Sub Tst()") Then Exit Function
If HasPfx(L, "Sub Z()") Then Exit Function
If HasPfx(L, "Sub Z()") Then Exit Function
If HasPfx(L, "Friend Sub Z()") Then Exit Function
LinIsTstSub = False
End Function
