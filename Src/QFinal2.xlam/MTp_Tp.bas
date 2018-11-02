Attribute VB_Name = "MTp_Tp"
Option Explicit

Function TpT1LyDic(A) As Dictionary

End Function

Sub TpBrkAsg(A$, OErLy$(), ORmkDic As Dictionary, Ny0, ParamArray OLyAp())
Dim O(), J%, U%
O = ClnBrk1(LyCln(SplitCrLf(A)), Ny0)
U = UB(O)
For J = 0 To U - 2
    OLyAp(J) = O(J)
Next
OErLy = O(U + 1)
'Set ORmkDic = O(U + 2)
End Sub

Function LyLnxAy(A$()) As Lnx()
Dim J&, O() As Lnx
If Sz(A) = 0 Then Exit Function
For J = 0 To UB(A)
    PushObj O, Lnx(J, A(J))
Next
LyLnxAy = O
End Function
