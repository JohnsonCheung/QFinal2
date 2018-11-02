Attribute VB_Name = "MIde___EnmAndTy"
Option Explicit
Type MthPmTy
    TyChr As String
    TyAsNm As String
    IsAy As Boolean
End Type

Type MthPm
    Nm As String
    IsOpt As Boolean
    IsPmAy As Boolean
    Ty As MthPmTy
    DftVal As String
End Type

Type MthSig
    HasRetVal As Boolean
    PmAy() As MthPm
    RetTy As MthPmTy
End Type
Type LCC
    Lno As Long
    C1 As Integer
    C2 As Integer
End Type
Type RRCC
    R1 As Long
    R2 As Long
    C1 As Integer
    C2 As Integer
End Type

Function RRCC(R1, R2, C1, C2) As RRCC
With RRCC
    .R1 = R1
    .R2 = R2
    .C1 = C1
    .C2 = C2
End With
End Function
Function LCC(Lno, C1, C2) As LCC
With LCC
    .Lno = Lno
    .C1 = C1
    .C2 = C2
End With
End Function
