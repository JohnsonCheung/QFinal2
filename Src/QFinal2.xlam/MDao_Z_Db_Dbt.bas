Attribute VB_Name = "MDao_Z_Db_Dbt"
Option Explicit
Const CMod$ = "DaoDbt"
Function X() As Sql_Shared
Static Y As Sql_Shared
If IsNothing(Y) Then Set Y = Sql_Shared
Set X = Y
End Function
Sub ZZ_DbtCrtDupKeyRecTbl()
TblDrp "#A #B"
DoCmd.RunSQL "Select Distinct Sku,BchNo,CLng(Rate) as RateRnd into [#A] from ZZ_DbtUpdSeq"
DbtCrtDupKeyRecTbl CurDb, "#A", "Sku BchNo", "#B"
TblBrw "#B"
Stop
TblDrp "#B"
End Sub
Function IsDbtExist(A As Database, T$) As Boolean
IsDbtExist = Not A.OpenRecordset("Select Name from MSysObjects where Type in (1,6) and Name='?'").EOF
End Function

Private Sub ZZ_DbtPk()
Dim A As Database
Set A = SampleDb_DutyPrepare
Dim Dr(), Dry(), T
For Each T In DbTny(A)
    Erase Dr
    Push Dr, T
    PushAy Dr, DbtPk(A, CStr(T))
    Push Dry, Dr
Next
DryBrw Dry
End Sub

Sub ZZ_DbtPutTp()

End Sub

Private Sub ZZ_DbtSpecReSeq()
DbtSpecReSeq CurDb, "ZZ_DbtUpdSeq", "Permit PermitD"
End Sub

Sub ZZ_DbtSpecReSeqFld()
DbtSpecReSeqFld CurDb, "ZZ_DbtUpdSeq", "Permit PermitD"
End Sub

Private Sub ZZ_DbtUpdSeq()
DoCmd.SetWarnings False
DoCmd.RunSQL "Select * into [#A] from ZZ_DbtUpdSeq order by Sku,PermitDate"
DoCmd.RunSQL "Update [#A] set BchRateSeq=0, Rate=Round(Rate,0)"
DbtUpdSeq CurDb, "#A", "BchRateSeq", "Sku", "Sku Rate"
TblOpn "#A"
Stop
DoCmd.RunSQL "Drop Table [#A]"
End Sub

Sub ZZ_DbtUpdToDteFld()
DoCmd.RunSQL "Select * into [#A] from ZZ_DbtUpdToDteFld order by Sku,PermitDate"
DbtUpdToDteFld CurDb, "#A", "PermitDateEnd", "Sku", "PermitDate"
Stop
TblDrp "#A"
End Sub

Sub ZZ_DbtWhDupKey()
TTDrp "#A #B"
DoCmd.RunSQL "Select Distinct Sku,BchNo,CLng(Rate) as RateRnd into [#A] from ZZ_DbtUpdSeq"
DbtWhDupKey CurDb, "#A", "Sku BchNo", "#B"
TTBrw "#B"
Stop
TTDrp "#B"
End Sub
Sub DbtSetFDes(A As Database, T, B As FDes)
DbtfDes(A, T, B.F) = B.Des
End Sub
Function DbtFFSql$(A As Database, T, FF)
DbtFFSql = QSel_FF_Fm(AyReOrdAy(DbtFny(A, T), CvNy(FF)), T)
End Function
Function DbtAddFd(A As Database, T, Fd As DAO.Fields) As DAO.Field2
A.TableDefs(T).Fields.Append Fd
Set DbtAddFd = Fd
End Function

Sub DbtAddFld(A As Database, T, F, Ty As DataTypeEnum, Optional Sz%, Optional Precious%)
If DbtHasFld(A, T, F) Then Exit Sub
Dim S$, SqlTy$
SqlTy = DaoTySqlTy(Ty, Sz, Precious)
S = FmtQQ("Alter Table [?] Add Column [?] ?", T, F, Ty)
A.Execute S
End Sub

Function DbtTim(A As Database, T) As Date
DbtTim = A.TableDefs(T).Properties("LastUpdated").Value
End Function
Sub DbtAddPfx(A As Database, T, Pfx)
DbtRen A, T, Pfx & T
End Sub

Sub DbtBrw(A As Database, T)
DtBrw DbtDt(A, T)
End Sub

Function DbtChkCol(A As Database, T, LnkColStr$) As String()
Dim Ay() As LnkCol, O$(), Fny$(), J%, Ty As DAO.DataTypeEnum, F$
Ay = LnkColStr_LnkColAy(LnkColStr)
Fny = LnkColAy_ExtNy(Ay)
O = DbtChkFny(A, T, Fny)
If Sz(O) > 0 Then DbtChkCol = O: Exit Function
For J = 0 To UB(Ay)
    F = Ay(J).Extnm
    Ty = Ay(J).Ty
    PushNonBlankStr O, DbtChkFldType(A, T, F, Ty)
Next
If Sz(0) > 0 Then
    PushMsgUnderLin O, "Some field has unexpected type"
    DbtChkCol = O
End If
End Function

Function DbtChkFldType$(A As Database, T, F, Ty As DAO.DataTypeEnum)
Dim ActTy As DAO.DataTypeEnum
ActTy = A.TableDefs(T).Fields(F).Type
If ActTy <> Ty Then
    DbtChkFldType = FmtQQ("Table[?] field[?] should have type[?], but now it has type[?]", T, F, DaoTyShtStr(Ty), DaoTyShtStr(ActTy))
End If
End Function

Function DbtChkFny(A As Database, T, ExpFny$()) As String()
Dim Miss$(), TFny$(), O$(), I
TFny = DbtFny(A, T)
Miss = AyMinus(ExpFny, TFny)
DbtChkFny = DbtMissFny_Er(A, T, Miss, TFny)
End Function

Function DbtColChk(A As Database, T, ColLnk$()) As String()
Dim O$(), ExpFny$()
ExpFny = ColLnkExpFny(ColLnk)
O = DbtExpFnyChk(A, T, ExpFny)
If Sz(O) > 0 Then DbtColChk = O: Exit Function
O = DbtTyChk(A, T, ColLnk)
If Sz(O) > 0 Then
    PushMsgUnderLin O, "Some field has unexpected type"
    'DbtTyChk = O
End If
End Function

Sub DbtCrtDupKeyRecTbl(A As Database, T, KeySsl$, TarTbl$)
Dim Ky$(), K$, Jn$, Tmp$, J%
Tmp = "##" & Format(Now, "HHMMSS")
Ky = SslSy(KeySsl)
K = JnComma(Ky)
For J = 0 To UB(Ky)
    Ky(J) = FmtQQ("x.?=a.?", Ky(J), Ky(J))
Next
Jn = Join(Ky, " and ")
A.Execute FmtQQ("Select Distinct ?,Count(*) as Cnt into [?] from [?] group by ? having Count(*)>1", K, Tmp, T, K)
A.Execute FmtQQ("Select x.* into [?] from [?] x inner join [?] a on ?", TarTbl, T, Tmp, Jn)
DbtDrp A, Tmp
End Sub

Sub DbtCrtFldLisTbl(A As Database, T, TarTbl$, KyFld$, Fld$, Sep$, Optional IsMem As Boolean, Optional LisFldNm0$)
Dim LisFldNm$: LisFldNm = DftStr(LisFldNm0, Fld & "Lis")
Dim RsFm As DAO.Recordset, LasK, K, RsTo As DAO.Recordset, Lis$
A.Execute FmtQQ("Select ? into [?] from [?] where False", KyFld, TarTbl, T)
A.Execute FmtQQ("Alter Table [?] add column ? ?", TarTbl, LisFldNm, IIf(IsMem, "Memo", "Text(255)"))
Set RsFm = DbqRs(A, FmtQQ("Select ?,? From [?] order by ?,?", KyFld, Fld, T, KyFld, Fld))
If Not RsAny(RsFm) Then Exit Sub
Set RsTo = DbtRs(A, TarTbl)
With RsFm
    LasK = .Fields(0).Value
    Lis = .Fields(1).Value
    .MoveNext
    While Not .EOF
        K = .Fields(0).Value
        If LasK = K Then
            Lis = Lis & Sep & .Fields(1).Value
        Else
            DrInsRs Array(LasK, Lis), RsTo
            LasK = K
            Lis = .Fields(1).Value
        End If
        .MoveNext
    Wend
End With
DrInsRs Array(LasK, Lis), RsTo
End Sub

Sub DbtCrtPk(A As Database, T)
Q = CrtPkSql(T): A.Execute Q
End Sub

Sub DbtCrtSk(A As Database, T, Fny0)
Q = FmtQQ("Create Unique Index SecondaryKey on ? (?)", T, JnComma(CvNy(Fny0))): A.Execute Q
End Sub

Function DbtCsv(A As Database, T) As String()
DbtCsv = RsCsvLy(DbtRs(A, T))
End Function

Property Let DbtDes(A As Database, T, Des$)
DbtPrp(A, T, C_Des) = Des
End Property

Property Get DbtDes$(A As Database, T)
DbtDes = DbtPrp(A, T, C_Des)
End Property

Function DbtDftFny(A As Database, T, Optional Fny0) As String()
If IsMissing(Fny0) Then
   DbtDftFny = DbtFny(A, T)
Else
   DbtDftFny = CvNy(Fny0)
End If
End Function

Sub DbtDrp(A As Database, T)
If DbHasTbl(A, T) Then A.Execute DrpTblSql(T)
End Sub

Sub DbtDrpFld(A As Database, T, Fny0)
Dim F
For Each F In AyNz(CvNy(Fny0))
    A.Execute DrpFldSql(T, F)
Next
End Sub

Function DbtDrs(A As Database, T) As Drs
Set DbtDrs = RsDrs(DbtRs(A, T))
End Function

Function DbtDry(A As Database, T) As Variant()
DbtDry = RsDry(DbtRs(A, T))
End Function

Function DbtDt(A As Database, T) As Dt
Dim Fny$(): Fny = DbtFny(A, T)
Dim Dry(): Dry = RsDry(DbtRs(A, T))
Set DbtDt = Dt(T, Fny, Dry)
End Function

Function DbtExpFnyChk(A As Database, T, ExpFny0) As String()
Dim Miss$(), TFny$()
TFny = DbtFny(A, T)
Miss = AyMinus(CvNy(ExpFny0), TFny)
DbtExpFnyChk = MissingFnyChk(Miss, TFny, A, T)
End Function

Function DbtfFd(A As Database, T, F) As DAO.Field2
Set DbtfFd = A.TableDefs(T).Fields(F)
End Function

Function DbtFFdInfDr(A As Database, T, F) As Variant()
Dim FF  As DAO.Field
Set FF = A.TableDefs(T).Fields(F)
With FF
    DbtFFdInfDr = Array(F, IIf(DbtfIsPk(A, T, F), "*", ""), DaoTyStr(.Type), .Size, .DefaultValue, .Required, FldDes(FF))
End With
End Function

Function DbtfIsPk(A As Database, T, F) As Boolean
DbtfIsPk = AyHas(DbtPk(A, T), F)
End Function

Function DbtfNxtId&(A As Database, T, Optional F)
Dim S$: S = FmtQQ("select Max(?) from ?", Dft(F, T), T)
DbtfNxtId = DbqV(A, S) + 1
End Function

Sub DbtfAddExpr(A As Database, T, F, Expr$, Optional Ty As DAO.DataTypeEnum = dbText, Optional TxtSz As Byte = 255)
A.TableDefs(T).Fields.Append NewFd(F, Ty, TxtSz, False, Expr)
End Sub

Function DbtfAyInto(A As DAO.Database, T, F, OInto)
Dim Rs As DAO.Recordset
Set Rs = DbtfRs(A, T, F)
DbtfAyInto = RsAyInto(Rs, OInto)
End Function

Sub DbtfChgDteToTxt(A As Database, T, F)
A.Execute FmtQQ("Alter Table [?] add column [###] text(12)", T)
A.Execute FmtQQ("Update [?] set [###] = Format([?],'YYYY-MM-DD')", T, F)
A.Execute FmtQQ("Alter Table [?] Drop Column [?]", T, F)
A.Execute FmtQQ("Alter Table [?] Add Column [?] text(12)", T, F)
A.Execute FmtQQ("Update [?] set [?] = [###]", T, F)
A.Execute FmtQQ("Alter Table [?] Drop Column [###]", T)
End Sub

Property Get DbtfDes$(A As Database, T, F)
DbtfDes = DbtfPrp(A, T, F, C_Des)
End Property

Property Let DbtfDes(A As Database, T, F, Des$)
DbtfPrp(A, T, F, C_Des) = Des
End Property

Function DbtfEleScl$(A As DAO.Database, T, F)
Dim Td As DAO.TableDef, Fd As DAO.Field2
Set Td = A.TableDefs(T)
Set Fd = Td.Fields(F)
DbtfEleScl = FdEleScl(Fd)
End Function

Function DbtfidRs(A As Database, T, F, Id&) As DAO.Recordset
Q = FmtQQ("Select ? From ? where ?=?", F, T, T, Id)
Set DbtfidRs = A.OpenRecordset(Q)
End Function

Property Get DbtfidV(A As Database, T, F, Id&)
DbtfidV = DbtfidRs(A, T, F, Id).Fields(0).Value
End Property

Property Let DbtfidV(A As Database, T, F, Id&, V)
With DbtfidRs(A, T, F, Id)
    .Edit
    .Fields(0).Value = V
    .Update
End With
End Property

Function DbtfIntAy(A As DAO.Database, T, F) As Integer()
Q = FmtQQ("Select [?] from [?]", F, T)
DbtfIntAy = DbqIntAy(A, Q)
End Function

Property Let Dbtfk0V(A As Database, T, F, K0, V)
Dim W$, Sk$(), Rs As DAO.Recordset, Vy
Vy = CvNy(K0)
Sk = DbtSk(A, T)
W = KyK0_BExpr(Sk, K0)
Q = FmtQQ("Select ?,? from [?] where ?", F, JnComma(Sk), T, W)
Set Rs = A.OpenRecordset(Q)
If RsNoRec(Rs) Then
    DrInsRs ItmAddAy(V, Vy), Rs
Else
    DrUpdRs ApAy(V), Rs
End If
End Property

Property Get Dbtfk0V(A As Database, T, F, K0)
Dim W$, Sk$(), Rs As DAO.Recordset
Sk = DbtSk(A, T)
If Sz(Sk) <> Sz(K0) Then
    Er "Dbtfk0V", "In [Db], [T] with [Sk] of [Sk-Sz] is given a [K0] of [K0-Sz] to value of [F], but the sz don't match", DbNm(A), T, JnSpc(Sk), Sz(Sk), K0, Sz(K0), F
End If
W = KyK0_BExpr(Sk, K0)
Q = FmtQQ("Select ? from [?] where ?", F, T, W)
Set Rs = A.OpenRecordset(Q)
If RsNoRec(Rs) Then Exit Property
Dbtfk0V = Nz(RsFldVal(Rs, F), Empty)
End Property

Function DbtFds(A As Database, T) As DAO.Fields
Set DbtFds = A.TableDefs(T).Fields
End Function

Function DbtFny(A As Database, T) As String()
DbtFny = ItrNy(A.TableDefs(T).Fields)
End Function

Function DbtFnyQuoted(A As Database, T) As String()
Dim O$()
O = DbtFny(A, T)
If DbtIsXls(A, T) Then O = AyQuoteSqBkt(O)
DbtFnyQuoted = O
End Function


Function DbtfRs(A As DAO.Database, T, F) As DAO.Recordset
Set DbtfRs = DbqRs(A, SelTFSql(T, F))
End Function

Function DbtfSy(A As DAO.Database, T, F) As String()
DbtfSy = DbtfAyInto(A, T, F, DbtfSy)
End Function

Function DbtFstFldNm$(A As Database, T)
DbtFstFldNm = A.TableDefs(T).Fields(0).Name
End Function

Function DbtSy(A As DAO.Database, T) As String()
DbtSy = DbtfSy(A, T, DbtFstFldNm(A, T))
End Function

Function DbtfTy(A As Database, T, F) As DAO.DataTypeEnum
DbtfTy = A.TableDefs(T).Fields(F).Type
End Function

Function DbtfTyStr$(A As Database, T, F)
DbtfTyStr = DaoTyStr(DbtfTy(A, T, F))
End Function

Function DbtfVal(A As Database, T, F)
DbtfVal = A.TableDefs(T).OpenRecordset.Fields(F).Value
End Function

Function DbtFxOfLnkTbl$(A As Database, T)
DbtFxOfLnkTbl = TakBet(A.TableDefs(T).Connect, "Database=", ";")
End Function

Function DbtHasFld(A As Database, T, F) As Boolean
Ass DbtIsExist(A, T)
DbtHasFld = TblHasFld(A.TableDefs(T), F)
End Function

Function DbtHasLnk(A As Database, T, S$, Cn$)
Dim I As DAO.TableDef
For Each I In A.TableDefs
    If I.Name = T Then
        If I.SourceTableName <> S Then Exit Function
        If EnsSfxSC(I.Connect) <> EnsSfxSC(Cn) Then Exit Function
        DbtHasLnk = True
        Exit Function
    End If
Next
End Function

Function DbtHasUniqIdx(A As Database, T) As Boolean
DbtHasUniqIdx = ItrHasPrpTrue(A.TableDefs(T).Indexes, "Unique")
End Function

Function DbtIdRs(A As Database, T, Id&) As DAO.Recordset
Q = FmtQQ("Select * From ? where ?=?", T, T, Id)
Set DbtIdRs = A.OpenRecordset(Q)
End Function

Function DbtIdx(A As Database, T, Idx) As DAO.Index
Dim I As DAO.Index
For Each I In A.TableDefs(T).Indexes
    If I.Name = Idx Then Set DbtIdx = I: Exit Function
Next
End Function

Sub DbtImp(A As Database, T, ColLnk$())
DbtDrp A, "#I" & Mid(T, 2)
Q = ColLnk_ImpSql(ColLnk, T)
A.Execute Q
End Sub

Sub DbtImpMap(A As Database, T, LnkColStr$, Optional WhBExpr$)
If FstChr(T) <> ">" Then
    Debug.Print "FstChr of T must be >"
    Stop
End If
'Assume [>?] T exist
'Create [#I?] T
Dim S$
S = LnkColStr_ImpSql(LnkColStr, T, WhBExpr)
DbtDrp A, "#I" & Mid(T, 2)
A.Execute S
End Sub

Sub DbttImp(A As Database, TT)
Dim Tny$(), J%, S$
Tny = CvNy(TT)
For J = 0 To UB(Tny)
    DbtDrp A, "#I" & Tny(J)
    S = FmtQQ("Select * into [#I?] from [?]", Tny(J), Tny(J))
    A.Execute S
Next
End Sub

Function DbtIsExist(A As Database, T) As Boolean
DbtIsExist = FbHasTbl(A.Name, T)
'DbtIsExist = Not A.OpenRecordset("Select Name from MSysObjects where Type in (1,6) and Name='?'").EOF
End Function

Function DbtIsFbLnk(A As Database, T) As Boolean
DbtIsFbLnk = HasPfx(DbtCnStr(A, T), ";Database=")
End Function

Function DbtIsFxLnk(A As Database, T) As Boolean
DbtIsFxLnk = HasPfx(DbtCnStr(A, T), "Excel")
End Function

Function DbtIsSys(A As Database, T) As Boolean
DbtIsSys = A.TableDefs(T).Attributes And DAO.TableDefAttributeEnum.dbSystemObject
End Function

Function DbtIsXls(A As Database, T) As Boolean
DbtIsXls = HasPfx(DbtCnStr(A, T), "Excel")
End Function

Sub Dbtk0Ins(A As Database, T, K0)
DrInsRs K0Vy(K0), DbtRs(A, T)
End Sub

Function K0Vy(K0)
Select Case True
Case IsStr(K0): K0Vy = SslSy(K0)
Case IsArray(K0): K0Vy = K0
Case Else: Er "K0Vy", "K0 should either be string or array, but now it has [typename]", TypeName(K0)
End Select
End Function

Function DbtSsk$(A As Database, T)
DbtSsk = AyFstEle(DbtSk(A, T))
End Function

Function DbtSkAvWhStr$(A As Database, T, SkAv())
DbtSkAvWhStr = X.WhFnyEqAy(DbtSk(A, T), SkAv)
End Function

Function DbtkIsExist(A As Database, T, K&) As Boolean
DbtkIsExist = Not RsAny(DbtkRs(A, T, K))
End Function

Function DbtkRs(A As Database, T, K&) As DAO.Recordset
Set DbtkRs = A.OpenRecordset(QSel_Fm(T, X.WhK(K, T)))
End Function

Function DbtkfV(A As Database, T, K&, F) ' K is Pk value
DbtkfV = DbqVal(A, QSel_FF_Fm(T, F, X.WhK(K, T)))
End Function


Function DbttLnkFb(A As Database, TT$, Fb$, Optional FbTny0) As String()
Dim Tny$(), FbTny$(), J%, T
Tny = CvNy(TT)
FbTny = CvNy(FbTny0)
    Select Case Sz(FbTny)
    Case Sz(Tny)
    Case 0:    FbTny = Tny
    Case Else: Er "DbttLnkFb", "[TT]-[Sz1] and [Fbtt]-[Sz2] are diff.  (@DbttLnkFb)", TT, Sz(Tny), FbTny, Sz(FbTny)
    End Select
Dim Cn$: Cn = FbDaoCnStr(Fb)
For Each T In Tny
    PushIAy DbttLnkFb, DbtLnk(A, CStr(T), FbTny(J), Cn)
    J = J + 1
Next
End Function

Function DbtXlsLnkInf(A As Database, T) As XlsLnkInf
Dim Cn$
Cn = DbtCnStr(A, T)
If Not HasPfx(Cn, "Excel") Then Exit Function
With DbtXlsLnkInf
    .IsXlsLnk = True
    .Fx = TakBefOrAll(TakAft(Cn, "DATABASE="), ";")
    .WsNm = A.TableDefs(T).SourceTableName
    If LasChr(.WsNm) <> "$" Then Stop
    .WsNm = RmvLasChr(.WsNm)
End With
End Function

Function DbtMissFny_Er(A As Database, T, MissFny$(), ExistingFny$()) As String()
Dim X As XlsLnkInf, O$(), I
If Sz(MissFny) = 0 Then Exit Function
X = DbtXlsLnkInf(A, T)
If X.IsXlsLnk Then
    Push O, "Excel File       : " & X.Fx
    Push O, "Worksheet        : " & X.WsNm
    PushUnderLin O
    For Each I In ExistingFny
        Push O, "Worksheet Column : " & QuoteSqBkt(CStr(I))
    Next
    PushUnderLin O
    For Each I In MissFny
        Push O, "Missing Column   : " & QuoteSqBkt(CStr(I))
    Next
    PushMsgUnderLinDbl O, "Columns are missing"
Else
    Push O, "Database : " & A.Name
    Push O, "Table    : " & T
    For Each I In MissFny
        Push O, "Field    : " & QuoteSqBkt(CStr(I))
    Next
    PushMsgUnderLinDbl O, "Above Fields are missing"
End If
DbtMissFny_Er = O
End Function

Function DbtMsgPk$(A As Database, T)
Dim S%, K$(), O
K = IdxFny(DbtPIdx(A, T))
S = Sz(K)
Select Case True
Case S = 0: O = FmtQQ("T[?] has no Pk", T)
Case S = 1:
    If K(0) <> T & "Id" Then
        O = FmtQQ("T[?] has 1-field-Pk of Fld[?].  It should be [?Id]", T, K(0), T)
    End If
Case Else
    O = FmtQQ("T[?] has primary key.  It should have single field and name eq to table, but now it has Pk[?]", T, JnSpc(K))
End Select
DbtMsgPk = O
End Function

Function DbtMsgSk$(A As Database, T)
Dim I As DAO.Index, NoSk As Boolean
Set I = ItrFstNm(A.TableDefs(T).Indexes, "SecondaryKey")
NoSk = IsNothing(I)
Select Case True
Case IsUniqIdx(I): DbtMsgSk = MsgLin("[T] in [Db] has Idx-SecondaryKey should be Unique", T, DbNm(A)): Exit Function
Case NoSk And DbtHasUniqIdx(A, T):
    If Not IsNothing(I) Then
        FunMsgBrw "DbtSkIdx", "[T] of [Db] does not have Idx-SecondaryKey, but it has [Idx] with unique.  This 'Idx' should be named as 'SecondaryKey'", T, DbNm(A), I.Name
        Exit Function
    End If
    Exit Function
End Select
If Not I.Unique Then FunMsgBrw "DbtSkIdx", "IdxNm-SecondaryKey of [T] of [Db] should unique"
End Function

Function DbtNCol&(A As Database, T)
DbtNCol = A.TableDefs(T).Fields.Count
End Function

Function DbtNRec&(A As Database, T, Optional WhBExpr$)
DbtNRec = DbqV(A, FmtQQ("Select Count(*) from [?]?", T, X.Wh(WhBExpr)))
End Function

Function DbtPIdx(A As Database, T) As DAO.Index
Set DbtPIdx = ItrFstPrpTrue(A.TableDefs(T).Indexes, "Primary")
End Function

Function DbtPk(A As Database, T) As String()
DbtPk = IdxFny(DbtPIdx(A, T))
End Function

Function DbtPkIxNm$(A As Database, T)
ObjNm (DbtPIdx(A, T))
Dim I As DAO.Index
For Each I In A.TableDefs(T).Indexes
    If I.Primary Then DbtPkIxNm = I.Name
Next
End Function

Function DbtPkNm(A As Database, T)
Dim I As DAO.Index
For Each I In A.TableDefs(T).Indexes
    If I.Primary Then DbtPkNm = I.Name
Next
End Function



Function DbtRecCnt&(A As Database, T)
DbtRecCnt = DbqV(A, FmtQQ("Select Count(*) from [?]", T))
End Function

Sub DbtRen(A As Database, T, ToTbl)
FbDb(A.Name).TableDefs(T).Name = ToTbl
End Sub

Sub DbtRenCol(A As Database, T, Fm, NewCol)
FbDb(A.Name).TableDefs(T).Fields(Fm).Name = NewCol
End Sub

Sub DbtSpecReSeq(A As Database, T, ReSeqSpec$)
DbtFnyReSeq A, T, ReSeqSpecFny(ReSeqSpec)
End Sub

Sub DbtFnyReSeq(A As Database, T, Fny$())
Dim F$(), J%, FF, Flds As DAO.Fields
Set Flds = A.TableDefs(T).Fields
F = AyReOrdAy(Fny, DbtFny(A, T))
For Each FF In F
    J = J + 1
    Flds(FF).OrdinalPosition = J
Next
End Sub

Sub DbtSpecReSeqFld(A As Database, T, ReSeqSpec$)
DbtSpecReSeqFldByFny A, T, ReSeqSpecFny(ReSeqSpec)
End Sub

Sub DbtSpecReSeqFldByFny(A As Database, T, Fny$())
Dim TFny$(), F$(), J%, FF
TFny = DbtFny(A, T)
If Sz(TFny) = Sz(Fny) Then
    F = Fny
Else
    F = AyAdd(Fny, AyMinus(TFny, Fny))
End If
For Each FF In F
    J = J + 1
    A.TableDefs(T).Fields(FF).OrdinalPosition = J
Next
End Sub

Sub DbtReStru(A As Database, Stru, F As Drs, E As Dictionary)
Dim DrpFld$(), NewFld$(), T$, FnyO$(), FnyN$()
T = LinT1(Stru)
FnyO = DbtFny(A, T)
FnyN = StruFny(Stru)

DrpFld = AyMinus(FnyO, FnyN)
NewFld = AyMinus(FnyN, FnyO)
DbtDrpFld A, T, DrpFld
DbtAddFnyStruBase A, T, NewFld, F, E
End Sub


Function DbtRs(A As Database, T) As DAO.Recordset
Set DbtRs = A.OpenRecordset(T)
End Function

Function DbtScly(A As Database, T) As String()
DbtScly = TdScly(A.TableDefs(T))
End Function

Function DbtSecIdx(A As Database, T)
Set DbtSecIdx = DbtIdx(A, T, "SecondaryKey")
End Function

Sub DbtFDesDicSet(A As Database, T, FDes As Dictionary)
Dim FF, F$
For Each FF In DbtFny(A, T)
    F = FF
    If FDes.Exists(F) Then
        DbtfDes(A, T, F) = FDes(F)
    End If
Next
End Sub

Function DbtSIdx(A As Database, T) As DAO.Index
Dim O As DAO.Index
Set O = ItrFstPrpEqV(A.TableDefs(T).Indexes, "Name", T)
If Not O.Unique Then
    Er "DbtSIdx", "[Tbl] has index of same name, but not unique in [Db]", T, DbNm(A)
End If
Set DbtSIdx = O
End Function

Function DbtSimTyAy(A As Database, T, Optional Fny0) As eSimTy()
Dim Fny$(): Fny = CvNy(Fny0)
Dim O() As eSimTy
   Dim U%
   ReDim O(U)
   Dim J%, F
   J = 0
   For Each F In Fny
       O(J) = DaoTySim(DbtfFd(A, T, CStr(F)).Type)
       J = J + 1
   Next
DbtSimTyAy = O
End Function

Function DbtSk(A As Database, T) As String()
DbtSk = IdxFny(DbtSkIdx(A, T))
End Function

Function DbtSkIdx(A As Database, T) As DAO.Index
Dim O As DAO.Index
Set O = DbtSecIdx(A, T)
If IsNothing(O) Then Exit Function
If Not O.Unique Then Er "DbtSkIdx", "[T] of [Db] has Idx-SecondaryKey.  It should be Unique", DbNm(A), T
If O.Primary Then Er "DbtSkIdx", "[T] of [Db] is Primary, but is has a name-SecondaryKey.", DbNm(A), T
Set DbtSkIdx = O
End Function

Function DbtReSeqSq(A As Database, T, ReSeqSpec$) As Variant()
Stop '
End Function

Function DbtSq(A As Database, T) As Variant()
Dim NR&, NC&, Rs As DAO.Recordset
Dim O(), J&
NR = DbtNRec(A, T)
NC = DbtNCol(A, T)
Set Rs = DbtRs(A, T)
ReDim O(1 To NR + 1, 1 To NC)
With Rs
    DrPutSq ItrNy(.Fields), O
    J = 2
    While Not .EOF
        RsPutSq Rs, O, J
        J = J + 1
        .MoveNext
    Wend
    .Close
End With
DbtSq = O
End Function

Function DbtSelSq(A As Database, T, ReSeqSpec$) As Variant()
Q = QSel_FF_Fm(ReSeqSpecFny(ReSeqSpec), T)
DbtSelSq = RsSq(DbqRs(A, Q))
End Function

Function DbtSrcTblNm$(A As Database, T)
DbtSrcTblNm = A.TableDefs(T).SourceTableName
End Function

Function DbtStru$(A As Database, T)
Const CSub = CMod & "DbtStru"
If Not DbHasTbl(A, T) Then FunMsgLyDmp CSub, "[Db] has not such [Tbl]", DbNm(A), T: Exit Function
Dim F$()
    F = DbtFny(A, T)
    If DbtIsXls(A, T) Then
        DbtStru = T & " " & JnSpc(AyQuoteSqBktIfNeed(F))
        Exit Function
    End If

Dim P$
    If AyHas(F, T & "Id") Then
        P = " *Id"
        F = AyMinus(F, Array(T & "Id"))
    End If
Dim S, R
    Dim J%, X
    S = DbtSk(A, T)
    R = AyMinus(F, S)
    If Sz(S) > 0 Then
        For Each X In S
            S(J) = Replace(X, T, "*")
            J = J + 1
        Next
        S = " " & JnSpc(AyQuoteSqBktIfNeed(S)) & " |"
    Else
        S = ""
    End If
    '
    J = 0
    For Each X In R
        R(J) = Replace(X, T, "*")
        J = J + 1
    Next
R = " " & JnSpc(AyQuoteSqBktIfNeed(R))
DbtStru = T & P & S & R
End Function

Sub DbttAddPfx(A As Database, TT0, Pfx)
AyDoAXB CvNy(TT0), "DbtAddPfx", A, Pfx
End Sub

Sub DbttBrw(A As Database, TT0)
AyDoPX CvNy(TT0), "DbtBrw", A
End Sub

Sub DbttCrtPk(A As Database, TT0)
AyDoPX CvNy(TT0), "DbtCrtPk", A
End Sub

Sub DbttDrp(A As Database, TT)
Dim T
For Each T In CvNy(TT)
    DbtDrp A, CStr(T)
Next
End Sub

Function DbttStru$(A As Database, TT)
Dim T, O$()
For Each T In AySrt(CvNy(TT))
    PushI O, DbtStru(A, CStr(T))
Next
DbttStru = JnCrLf(AyAlign1T(O))
End Function


Function DbtTyChk(A As Database, T, ColLnk$()) As String()
Dim ActTy As DAO.DataTypeEnum, TyAy() As DAO.DataTypeEnum
'ActTy = A.TableDefs(T).Fields(F).Type
'TyAy = SCShtTy_TyAy(SCShtTy)
If Not AyHas(TyAy, ActTy) Then
    Dim S1$, S2$
'    S1 = DaoTyAy_SCShtTy(TyAy)
    S2 = DaoTyShtStr(ActTy)
    'DbtfTyMsg = FmtQQ("Table[?] field[?] has type[?].  It should be type[?].", T, F, S1, S2)
End If
End Function

Sub DbtUpdIdFld(A As Database, T, Fld)
Dim D As New Dictionary, J&, Rs As DAO.Recordset, Id$, V
Id = Fld & "Id"
Set Rs = DbqRs(A, FmtQQ("Select ?,?Id from [?]", Fld, Fld, T))
With Rs
    While Not .EOF
        .Edit
        V = .Fields(0).Value
        If D.Exists(V) Then
            .Fields(1).Value = D(V)
        Else
            .Fields(1).Value = J
            D.Add V, J
            J = J + 1
        End If
        .Update
        .MoveNext
    Wend
End With
End Sub

Sub DbtUpdSeq(A As Database, T, SeqFldNm$, Optional RestFny0, Optional IncFny0)
'Assume T is sorted
'
'Update A->T->SeqFldNm using RestFny0,IncFny0, assume the table has been sorted
'Update A->T->SeqFldNm using OrdFny0, RestFny0,IncFny0
Dim RestFny$(), IncFny$(), Sql$
Dim LasRestVy(), LasIncVy(), Seq&, OrdS$, Rs As DAO.Recordset
'OrdFny RestAy IncAy Sql
RestFny = CvNy(RestFny0)
IncFny = CvNy(IncFny0)
If Sz(RestFny) = 0 And Sz(IncFny) = 0 Then
    With A.OpenRecordset(T)
        Seq = 1
        While Not .EOF
            .Edit
            .Fields(SeqFldNm) = Seq
            Seq = Seq + 1
            .Update
            .MoveNext
        Wend
        .Close
    End With
    Exit Sub
End If
'--
Set Rs = A.OpenRecordset(T) ', RecordOpenOptionsEnum.dbOpenForwardOnly, dbForwardOnly)
With Rs
    While Not .EOF
        If RsIsBrk(Rs, RestFny, LasRestVy) Then
            Seq = 1
            LasRestVy = RsVy(Rs, RestFny)
            LasIncVy = RsVy(Rs, IncFny)
        Else
            If RsIsBrk(Rs, IncFny, LasIncVy) Then
                Seq = Seq + 1
                LasIncVy = RsVy(Rs, IncFny)
            End If
        End If
        .Edit
        .Fields(SeqFldNm).Value = Seq
        .Update
        .MoveNext
    Wend
End With
End Sub

Sub DbtUpdToDteFld(A As Database, T, ToDteFld$, KeyFld$, FmDteFld$)
Dim ToDte() As Date, J&
ToDte = DbtUpdToDteFld__1(A, T, KeyFld, FmDteFld)
With DbtRs(A, T)
    While Not .EOF
        .Edit
        .Fields(ToDteFld).Value = ToDte(J): J = J + 1
        .Update
        .MoveNext
    Wend
    .Close
End With
End Sub

Function DbtUpdToDteFld__1(A As Database, T, KeyFld$, FmDteFld$) As Date()
Dim K$(), FmDte() As Date, ToDte() As Date, J&, CurKey$, NxtKey$, NxtFmDte As Date
With DbtRs(A, T)
    While Not .EOF
        Push FmDte, .Fields(FmDteFld).Value
        Push K, .Fields(KeyFld).Value
        .MoveNext
    Wend
End With
Dim U&
U = UB(K)
ReDim ToDte(U)
For J = 0 To U - 1
    CurKey = K(J)
    NxtKey = K(J + 1)
    NxtFmDte = FmDte(J + 1)
    If CurKey = NxtKey Then
        ToDte(J) = DateAdd("D", -1, NxtFmDte)
    Else
        ToDte(J) = DateSerial(2099, 12, 31)
    End If
Next
ToDte(U) = DateSerial(2099, 12, 31)
DbtUpdToDteFld__1 = ToDte
End Function


Sub DbtWhDupKey(A As Database, T, KK, TarTbl$)
Dim Ky$(), K$, Jn$, Tmp$, J%
Ky = SslSy(KK)
Tmp = "##" & TmpNm
K = JnComma(Ky)
For J = 0 To UB(Ky)
    Ky(J) = FmtQQ("x.?=a.?", Ky(J), Ky(J))
Next
Jn = Join(Ky, " and ")
A.Execute FmtQQ("Select Distinct ?,Count(*) as Cnt into [?] from [?] group by ? having Count(*)>1", K, Tmp, T, K)
A.Execute FmtQQ("Select x.* into [?] from [?] x inner join [?] a on ?", TarTbl, T, Tmp, Jn)
DbtDrp A, Tmp
End Sub

Sub DbtFFIntoAyAp(A As Database, T, FF, ParamArray OAyAp())
Dim Drs As Drs
'Set Drs = DbqDrs(MthDb, "Select MchStr,MchTy,MdNm From MthMch")
Dim F, J%
For Each F In CvNy(FF)
    OAyAp(J) = DrsColInto(Drs, F, OAyAp(J))
    J = J + 1
Next
End Sub

Function DbtFFDry(A As Database, T, FF) As Variant()
DbtFFDry = DbqDry(A, QSel_FF_Fm(FF, T))
End Function

Function DbtFFDrs(A As Database, T, FF) As Drs
Set DbtFFDrs = DbqDrs(A, QSel_FF_Fm(FF, T))
End Function
