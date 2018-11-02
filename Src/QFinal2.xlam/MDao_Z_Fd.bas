Attribute VB_Name = "MDao_Z_Fd"
Option Explicit
Function FdClone(A As DAO.Field2, FldNm) As DAO.Field2
Set FdClone = New DAO.Field
With FdClone
    .Name = FldNm
    .Type = A.Type
    .AllowZeroLength = A.AllowZeroLength
    .Attributes = A.Attributes
    .DefaultValue = A.DefaultValue
    .Expression = A.Expression
    .Required = A.Required
    .ValidationRule = A.ValidationRule
    .ValidationText = A.ValidationText
End With
End Function

Function FdDes$(A As DAO.Field)
If PrpHas(A.Properties, C_Des) Then FdDes = A.Properties(C_Des)
End Function

Function FdEleScl$(A As DAO.Field2)
Dim Rq$, Ty$, TxtSz$, AlwZ$, Rul$, Dft$, VTxt$, Expr$, Des$
Des = AddLbl(FdDes(A), "Des")
Rq = BoolTxt(A.Required, "Req")
AlwZ = BoolTxt(A.AllowZeroLength, "AlwZ")
Ty = DaoTyShtStr(A.Type)
If A.Type = DAO.DataTypeEnum.dbText Then TxtSz = BoolTxt(A.Type = dbText, "TxXTSz=" & A.Size)
Rul = AddLbl(A.ValidationText, "VTxt")
VTxt = AddLbl(A.ValidationRule, "VRul")
Expr = AddLbl(A.Expression, "Expr")
Dft = AddLbl(A.DefaultValue, "Dft")
FdEleScl = ApScl(Ty, TxtSz, Rq, AlwZ, Rul, VTxt, Dft, Expr)
End Function

Function FdScl$(A As DAO.Field2)
FdScl = A.Name & ";" & FdEleScl(A)
End Function

Function FdSclFd(ByVal A$) As DAO.Field2
Const CSub$ = "FdScl_Fd"
Dim J%, F$, L$, T$, Ay$(), Sz%, Des$, Rq As Boolean, Ty As DAO.DataTypeEnum, AlwZ As Boolean, Dft$, VRul$, VTxt$, Expr$, Er$()
If A = "" Then Exit Function
F = SclShf(A)
T = SclShf(A)
Ty = DaoShtTyStrTy(T)
SclAsg A, VdtEleSclNmSsl, Rq, AlwZ, Sz, Dft, VRul, VTxt, Des, Expr
Dim O As New DAO.Field
With O
    .Name = F
    .DefaultValue = Dft
    .Required = Rq
    .Type = Ty
    If Ty = DAO.DataTypeEnum.dbText Then
        .Size = Sz
        .AllowZeroLength = AlwZ
    End If
    .ValidationRule = VRul
    .ValidationText = VTxt
End With
Set FdSclFd = O
End Function

Function FdsCsv$(A As DAO.Fields)
FdsCsv = AyCsv(ItrVy(A))
End Function

Function FdsDr(A As DAO.Fields) As Variant()
Dim F As DAO.Field
For Each F In A
    PushI FdsDr, F.Value
Next
End Function

Function FdsFny(A As Fields) As String()
FdsFny = ItrNy(A)
End Function

Function FdsHasFld(A As DAO.Fields, F) As Boolean
FdsHasFld = ItrHasNm(A, F)
End Function

Function FdsKyDr(A As DAO.Fields, Ky0) As Variant()
Dim O(), K
For Each K In CvNy(Ky0)
    Push FdsKyDr, A(K).Value
Next
FdsKyDr = O
End Function

Sub FdsPutSq(A As DAO.Fields, Sq, Optional R& = 1, Optional NoTxtSngQ As Boolean)
Dim F As DAO.Field, J%
If NoTxtSngQ Then
    For Each F In A
        J = J + 1
        Sq(R, J) = F.Value
    Next
    Exit Sub
End If
For Each F In A
    J = J + 1
    If F.Type = DAO.DataTypeEnum.dbText Then
        Sq(R, J) = "'" & F.Value
    Else
        Sq(R, J) = F.Value
    End If
Next
End Sub

Sub FdsPutSq1(A As DAO.Fields, Sq, Optional R& = 1, Optional NoTxtSngQ As Boolean)
DrPutSq FdsDr(A), Sq, R, NoTxtSngQ
End Sub

Function FdsqlTy$(A As DAO.Field2)
Stop '
End Function

Function FdStr$(A As DAO.Field2)
Dim D$, R$, Z$, VTxt$, VRul, E$, S$
If A.Type = DAO.DataTypeEnum.dbText Then S = " TxtSz=" & A.Size
If A.DefaultValue <> "" Then D = " " & QuoteSqBktIfNeed("Dft=" & A.DefaultValue)
If A.Required Then R = " Req"
If A.AllowZeroLength Then Z = " AlwZLen"
If A.Expression <> "" Then E = " " & QuoteSqBktIfNeed("Expr=" & A.Expression)
If A.ValidationRule <> "" Then VRul = " " & QuoteSqBktIfNeed("VRul=" & A.ValidationRule)
If A.ValidationText <> "" Then VRul = " " & QuoteSqBktIfNeed("VTxt=" & A.ValidationText)
FdStr = A.Name & " " & DaoTyShtStr(A.Type) & R & Z & S & VTxt & VRul & D & E
End Function

Function FdsVy(A As DAO.Fields, Optional Ky0) As Variant()
Select Case True
Case IsMissing(Ky0)
    FdsVy = ItrVy(A)
Case IsStr(Ky0)
    FdsVy = FdsVyByKy(A, SslSy(Ky0))
Case IsSy(Ky0)
    FdsVy = FdsVyByKy(A, CvSy(Ky0))
Case Else
    Stop
End Select
End Function

Function FdsVyByKy(A As DAO.Fields, Ky$()) As Variant()
Dim O(), J%, K
If Sz(Ky) = 0 Then
    FdsVyByKy = ItrVy(A)
    Exit Function
End If
ReDim O(UB(Ky))
For Each K In Ky
    O(J) = A(K).Value
    J = J + 1
Next
FdsVyByKy = O
End Function

Function Fld2Lines$(A As DAO.Field2)
Dim O$, M$, Off&
X:
M = A.GetChunk(Off, 1024)
O = O & M
If Len(M) = 1024 Then
    Off = Off + 1024
    GoTo X
End If
Fld2Lines = O
End Function

Function FldDes$(A As DAO.Field)
FldDes = PrpVal(A.Properties, "Description")
End Function

Function FldInfDryFny() As String()
FldInfDryFny = SplitSpc("Fld Pk Ty Sz Dft Req Des")
End Function

Private Function FldInpy(NoT1$()) As String()
FldInpy = AyTakT1(LyFld(NoT1))
End Function

Function FldSqlTy$(Fld, F As Drs, E As Dictionary)
FldSqlTy = FdsqlTy(LookupFd(F, "", F, E))
End Function

Function FldVal(A As DAO.Field)
Asg A.Value, FldVal
End Function

Private Sub Z_FldsDr()
Dim Rs As DAO.Recordset, Dry()
'Set Rs = CurDb.OpenRecordset("Select * from Att")
With Rs
    While Not .EOF
        Push Dry, RsDr(Rs)
        .MoveNext
    Wend
    .Close
End With
End Sub

Function IsEqFd(A As DAO.Field2, B As DAO.Field2) As Boolean
With A
    If .Name <> B.Name Then Exit Function
    If .Type <> B.Type Then Exit Function
    If .Required <> B.Required Then Exit Function
    If .AllowZeroLength <> B.AllowZeroLength Then Exit Function
    If .DefaultValue <> B.DefaultValue Then Exit Function
    If .ValidationRule <> B.ValidationRule Then Exit Function
    If .ValidationText <> B.ValidationText Then Exit Function
    If .Expression <> B.Expression Then Exit Function
    If .Attributes <> B.Attributes Then Exit Function
    If .Size <> B.Size Then Exit Function
End With
IsEqFd = True
End Function
Sub ZZ_FldsVy()
Dim Rs As DAO.Recordset, Vy()
'Set Rs = CurDb.OpenRecordset("Select * from SkuB")
With Rs
    While Not .EOF
        Vy = RsVy(Rs)
        Debug.Print JnComma(Vy)
        .MoveNext
    Wend
    .Close
End With
End Sub
