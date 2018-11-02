Attribute VB_Name = "MDao_Schm_Std"
Option Explicit
Function StdCrtDteFd(Fld) As DAO.Field2
Set StdCrtDteFd = StdFd(Fld, dbDate, True)
StdCrtDteFd.DefaultValue = "Now"
End Function

Function StdCurFd(Fld) As DAO.Field2
Set StdCurFd = StdFd(Fld, dbCurrency, True)
StdCurFd.DefaultValue = 0
End Function

Function StdDteFd(Fld) As DAO.Field2
Set StdDteFd = StdFd(Fld, dbDate)
End Function

Function StdEleFd(Ele, Fld) As DAO.Field2
Dim O As DAO.Field2
Set O = StdEleTnnnFd(Ele, Fld): If Not IsNothing(O) Then Set StdEleFd = O: Exit Function
Select Case Ele
Case "Nm": Set StdEleFd = StdNmFd(Fld)
Case "Amt": Set StdEleFd = StdFd(Fld, dbCurrency, True): StdEleFd.DefaultValue = 0
Case "Txt": Set StdEleFd = StdFd(Fld, dbText, True): StdEleFd.DefaultValue = """""": StdEleFd.AllowZeroLength = True
Case "Cur": Set StdEleFd = StdFd(Fld, dbCurrency, True): StdEleFd.DefaultValue = 0
Case "Dte": Set StdEleFd = StdFd(Fld, dbDate, False)
Case "Int": Set StdEleFd = StdFd(Fld, dbInteger, True): StdEleFd.DefaultValue = 0
Case "Lng": Set StdEleFd = StdFd(Fld, dbLong, True): StdEleFd.DefaultValue = 0
Case "Dbl": Set StdEleFd = StdFd(Fld, dbDouble, True): StdEleFd.DefaultValue = 0
Case "Sng": Set StdEleFd = StdFd(Fld, dbSingle, True): StdEleFd.DefaultValue = 0
Case "Lgc": Set StdEleFd = StdFd(Fld, dbBoolean, True): StdEleFd.DefaultValue = 0
Case "Mem": Set StdEleFd = StdFd(Fld, dbMemo, True): StdEleFd.DefaultValue = """""": StdEleFd.AllowZeroLength = True
End Select
End Function

Function StdEleTnnnFd(Ele, Fld) As DAO.Field2
If Left(Ele, 1) <> "T" Then Exit Function
Dim A$
A = Mid(Ele, 2)
If CStr(Val(A)) <> A Then Exit Function
Set StdEleTnnnFd = StdFd(Fld, dbText, True)
With StdEleTnnnFd
    .Size = A
    .DefaultValue = """"""
    .AllowZeroLength = True
End With
End Function

Function StdFd(Fld, Ty As DAO.DataTypeEnum, Optional Req As Boolean) As DAO.Field2
Set StdFd = New DAO.Field
With StdFd
    .Name = Fld
    .Type = Ty
    .Size = 255
    .Required = Req
End With
End Function

Function StdFldFd(Fld, Tbl) As DAO.Field2
Dim R2$, R3$: R2 = Right(Fld, 2): R3 = Right(Fld, 3)
Select Case True
Case Fld = "CrtDte":   Set StdFldFd = StdCrtDteFd(Fld)
Case Tbl & "Id" = Fld: Set StdFldFd = StdPkFd(Fld)
Case R2 = "Id":        Set StdFldFd = StdIdFd(Fld)
Case R2 = "Ty":        Set StdFldFd = StdTyFd(Fld)
Case R2 = "Nm":        Set StdFldFd = StdNmFd(Fld)
Case R3 = "Dte":       Set StdFldFd = StdDteFd(Fld)
Case R3 = "Amt":       Set StdFldFd = StdCurFd(Fld)
End Select
End Function

Function StdIdFd(Fld) As DAO.Field2
Set StdIdFd = StdFd(Fld, dbLong, True)
End Function

Function StdNmFd(Fld) As DAO.Field2
Set StdNmFd = StdTxtFd(Fld, True, 50, False)
End Function

Function StdPkFd(Fld) As DAO.Field2
Set StdPkFd = StdFd(Fld, dbLong, True)
StdPkFd.Attributes = DAO.FieldAttributeEnum.dbAutoIncrField
End Function

Function StdTxtFd(Fld, Sz%, Optional Req As Boolean, Optional AlwZLen As Boolean) As DAO.Field2
Set StdTxtFd = StdFd(Fld, dbText, Req)
End Function

Function StdTyFd(Fld) As DAO.Field2
Set StdTyFd = StdTxtFd(Fld, 20, Req:=True, AlwZLen:=False)
End Function

Function IsStdFld(Fld) As Boolean
IsStdFld = True
If Fld = "CrtDte" Then Exit Function
If AyHas(SslSy("Id Ty Nm"), Right(Fld, 2)) Then Exit Function
If AyHas(SslSy("Dte Amt"), Right(Fld, 3)) Then Exit Function
IsStdFld = False
End Function

Function IsStdEle(Ele) As Boolean
Stop '
End Function

