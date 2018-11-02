Attribute VB_Name = "MVb_Ay_Cnt"
Option Explicit

Function AyCntDry(A) As Variant()
AyCntDry = DicDry(AyCntDic(A))
End Function

Private Sub Z_AyCntDry()
Dim A$()
A = SplitSpc("a a a b c b")
Ept = Array(Array("a", 3), Array("b", 2), Array("c", 1))
GoSub Tst
Exit Sub
Tst:
    Act = AyCntDry(A)
    Ass IsEqAy(Act, Ept)
    Return
End Sub

Function AyCntDic(A) As Dictionary
Dim O As New Dictionary, I
For Each I In AyNz(A)
    If O.Exists(I) Then
        O(I) = O(I) + 1
    Else
        O.Add I, 1
    End If
Next
Set AyCntDic = O
End Function
