Attribute VB_Name = "MDao_Z_Fd_New"
Option Explicit

Function NewFd_zFdScl(FdScl$) As DAO.Field2
Set NewFd_zFdScl = FdSclFd(FdScl)
End Function

Function NewFd_zFk(F) As DAO.Field2
Dim O As New DAO.Field
With O
    .Name = F
    .Type = dbLong
End With
Set NewFd_zFk = O
End Function

Function NewFd_zId(F) As DAO.Field2
Dim O As New DAO.Field
With O
    .Name = F
    .Type = dbLong
    .Attributes = DAO.FieldAttributeEnum.dbAutoIncrField
    .Required = True
End With
Set NewFd_zId = O
End Function

Function NewFd(F, Optional Ty As DAO.DataTypeEnum = dbText, Optional TxtSz As Byte = 255, Optional ZLen As Boolean, Optional Expr$, Optional Dft$, Optional Req As Boolean, Optional VRul$, Optional VTxt$) As DAO.Field2
Dim O As New DAO.Field
With O
    .Name = F
    .Required = Req
    If Ty <> 0 Then .Type = Ty
    If Ty = dbText Then
        .Size = TxtSz
        .AllowZeroLength = ZLen
    End If
    If Expr <> "" Then
        CvFd2(O).Expression = Expr
    End If
    O.DefaultValue = Dft
End With
Set NewFd = O
End Function
