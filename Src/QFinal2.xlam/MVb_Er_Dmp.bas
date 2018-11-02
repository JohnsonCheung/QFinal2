Attribute VB_Name = "MVb_Er_Dmp"
Option Explicit
Sub D(A)
Select Case True
Case IsArray(A): AyDmp A
Case IsDic(A):   DicDmp CvDic(A)
Case Else: Debug.Print A
End Select
End Sub

Sub DmpTy(A)
Debug.Print TypeName(A)
End Sub

Sub AyDmp(A, Optional WithIx As Boolean)
If Sz(A) = 0 Then Exit Sub
Dim I
If WithIx Then
    Dim J&
    For Each I In A
        Debug.Print J; ": "; I
        J = J + 1
    Next
Else
    For Each I In A
        Debug.Print I
    Next
End If
End Sub


Sub Chk(Check$())
If Sz(Check) = 0 Then Exit Sub
AyBrw Check
Stop
End Sub
Sub ChkEq(A, B)
If Not IsEq(A, B) Then
    Debug.Print "["; A; "] [" & TypeName(A) & "]"
    Debug.Print "["; B; "] [" & TypeName(A) & "]"
    Stop
End If
End Sub
