Attribute VB_Name = "MDao_Schm_Ele"
Option Explicit

Function EleDicFmt(A As Dictionary) As String()
If IsNothing(A) Then PushI EleDicFmt, "*Nothing": Exit Function
Dim K
For Each K In A.Keys
    PushI EleDicFmt, K & " " & CvFd(FdStr(A(K)))
Next
End Function

Function EleStrFd(A) As DAO.Field2
Dim TyStr$, R As Boolean, Z As Boolean, D$, VTxt$, VRul$, S$, X$
Dim L$: L = A
Dim Ay$()
Ay = ShfVal(L, EleLblss)
AyAsg Ay, _
    TyStr, R, Z, D, VTxt, VRul, S, X
Set EleStrFd = New DAO.Field
With EleStrFd
    .Type = DaoShtTyStrTy(TyStr)
    .Required = R
    If .Type = dbText Then .AllowZeroLength = Z
    .DefaultValue = D
    .ValidationText = VTxt
    .ValidationRule = VRul
    .Size = Val(S)
    .Expression = X
End With
End Function

Private Sub Z_EleStrFd()
Dim A$, Act As DAO.Field2, Ept As DAO.Field2
A = "Int Req AlwZLen Dft=ABC TxtSz=10"
Set Ept = New DAO.Field
With Ept
    .Type = DAO.DataTypeEnum.dbInteger
'    .AllowZeroLength = True
    .DefaultValue = "ABC"
    .Required = True
    .Size = 10
End With
GoSub Tst
Exit Sub
Tst:
    Set Act = EleStrFd(A)
    If Not IsEqFd(Act, Ept) Then Stop
    Return
End Sub


