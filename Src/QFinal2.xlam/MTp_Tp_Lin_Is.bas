Attribute VB_Name = "MTp_Tp_Lin_Is"
Option Explicit
Function LinIsRmkLin(A) As Boolean

End Function

Function LinIsTpRmkLin(A$) As Boolean
Dim L$: L = LTrim(A)
If L <> "" Then
    If HasPfx(L, "--") Then
        LinIsTpRmkLin = True
    End If
End If
End Function
