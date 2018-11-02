Attribute VB_Name = "MIde_Ty_Mth_ShtTy"
Option Explicit

Function MthTy_IsVdt(A) As Boolean
MthTy_IsVdt = AyHas(MthTyAy, A)
End Function

Function MthShtTyKd$(MthShtTy)
Select Case MthShtTy
Case "Fun", "Sub": MthShtTyKd = MthShtTy
Case "Get", "Let", "Set": MthShtTyKd = "Prp"
End Select
End Function



Function IsMthTy(A) As Boolean
IsMthTy = AyHas(MthTyAy, A)
End Function


Function IsMdy(A) As Boolean
IsMdy = AyHas(MdyAy, A)
End Function



Function MthKd$(MthTy$)
Select Case MthTy
Case "Function": MthKd = "Fun"
Case "Sub": MthKd = "Sub"
Case "Property Get", "Property Get", "Property Let": MthKd = "Prp"
End Select
End Function
