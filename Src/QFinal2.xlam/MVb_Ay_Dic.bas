Attribute VB_Name = "MVb_Ay_Dic"
Option Explicit
Function AyIxDic(A) As Dictionary
Dim O As New Dictionary, J&
For J = 0 To UB(A)
    O.Add A(J), J
Next
Set AyIxDic = O
End Function

Function AyDic(A, Optional V) As Dictionary
Set AyDic = New Dictionary
Dim I
For Each I In AyNz(A)
    AyDic.Add I, V
Next
End Function

Function AyDistIdCntDic(A) As Dictionary
'Type DistIdCntDic = Map Val [Id,Cnt]
Dim X, O As New Dictionary, J&, IdCnt()
For Each X In AyNz(A)
    If Not O.Exists(X) Then
        O.Add X, Array(J, 1)
        J = J + 1
    Else
        IdCnt = O(X)
        O(X) = Array(IdCnt(0), IdCnt(1) + 1)
    End If
Next
Set AyDistIdCntDic = O
End Function
