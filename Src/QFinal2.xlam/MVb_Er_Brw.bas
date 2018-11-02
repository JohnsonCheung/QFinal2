Attribute VB_Name = "MVb_Er_Brw"
Option Explicit

Function AyBrwEr(A, Optional Msg$) As Boolean
If Sz(A) = 0 Then Exit Function
Dim O
If Msg = "" Then
    O = A
Else
    O = AyInsAy(A, Array(Msg, UnderLin(Msg)))
End If
AyBrw O
AyBrwEr = True
End Function

Sub AyBrwThw(A, Optional Msg$)
If AyBrwEr(A, Msg) Then Err.Raise -1
End Sub
