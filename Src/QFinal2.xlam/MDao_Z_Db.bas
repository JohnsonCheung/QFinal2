Attribute VB_Name = "MDao_Z_Db"
Option Explicit
Function DbAddTd(A As Database, Td As DAO.TableDef) As DAO.TableDef
A.TableDefs.Append Td
Set DbAddTd = Td
End Function

Sub DbAddTmpTbl(A As Database)
DbAddTd CurDb, TmpTd
End Sub

Sub DbAppTd(A As Database, Td As DAO.TableDef)
A.TableDefs.Append Td
End Sub

Function DbChk(A As Database) As String()
Dim T$()
T = AySrt(DbTny(A))
DbChk = AyAlign1T(AyAddAp(DbChkPk(A, T), DbChkSk(A, T)))
End Function

Function DbChkPk(A As Database, Tny$()) As String()
DbChkPk = AyRmvEmp(AyMapPXSy(Tny, "DbtMsgPk", A))
End Function

Function DbChkSk(A As Database, Tny$()) As String()
End Function

Function DbCnSy(A As Database) As String()
Dim T$(), S()
T = AyQuoteSqBkt(DbTny(A))
S = AyMapPX(T, "DbtCnStr", A)
DbCnSy = AyabNonEmpBLy(T, S)
End Function

Sub DbCrtQry(A As Database, Q, Sql)
If Not DbHasQry(A, Q) Then
    Dim QQ As New QueryDef
    QQ.Sql = Sql
    QQ.Name = Q
    A.QueryDefs.Append QQ
Else
    A.QueryDefs(Q).Sql = Sql
End If
End Sub

Sub DbCrtResTbl(A As Database)
DbtDrp A, "Res"
DoCmd.RunSQL "Create Table Res (ResNm Text(50), Att Attachment)"
End Sub

Sub DbCrtSpecTbl(A As Database)
DbSchmCrt A, SpecSchmLines
End Sub

Sub DbCrtTbl(A As Database, T$, FldDclAy)
A.Execute FmtQQ("Create Table [?] (?)", T, JnComma(FldDclAy))
End Sub

Function DbDesy(A As Database) As String()
Dim T$(), D$()
T = DbTny(A)
DbDesy = AyRmvEmp(AyMapPXSy(T, "DbtTblDes", A))
End Function

Sub DbDrpAllTmpTbl(A As Database)
DbttDrp A, DbTmpTny(A)
End Sub

Sub DbDrpQry(A As Database, Q)
If DbHasQry(A, Q) Then A.QueryDefs.Delete Q
End Sub

Function DbDrsNormSqy(A As Database, B As Drs, Tny$()) As String()

End Function

Function DbDs(A As Database, Tny0, Optional DsNm$ = "Ds") As Ds
Dim DtAy1() As Dt
    Dim U%, Tny$()
    Tny = CvNy(Tny0)
    U = UB(Tny)
    ReDim DtAy(U)
    Dim J%
    For J = 0 To U
        Set DtAy(J) = DbtDt(A, Tny(J))
    Next
Set DbDs = Ds(DtAy1, DftDbNm(DsNm, A))
End Function

Private Sub Z_DbDs()
Dim Ds As Ds
Ds = DbDs(CurDb, "Permit PermitD")
Stop
End Sub

Sub DbEnsSpecTbl(A As Database)
If Not DbHasTbl(A, "Spec") Then DbCrtSpecTbl A
End Sub

Sub DbEnsTmp1Tbl(A As Database)
If DbHasTbl(A, "Tmp1") Then Exit Sub
DbqRun A, "Create Table Tmp1 (AA Int, BB Text 10)"
End Sub

Function DbHasQry(A As Database, Q) As Boolean
DbHasQry = DbqAny(A, FmtQQ("Select * from MSysObjects where Name='?' and Type=5", Q))
End Function

Function DbHasTbl(A As Database, T) As Boolean
DbHasTbl = CatHasTbl(FbCat(A.Name), T)
'DbHasTbl = DbqAny(A, FmtQQ("Select * from MSysObjects where Name='?' and Type in (1,6)", T))
End Function

Sub DbImpSpec(A As Database, Spnm)
Const CSub$ = "DbImpSpec"
Dim Ft$
    Ft = SpnmFt(Spnm)
    
Dim NoCur As Boolean
Dim NoLas As Boolean
Dim CurOld As Boolean
Dim CurNew As Boolean
Dim SamTim As Boolean
Dim DifSz As Boolean
Dim SamSz As Boolean
Dim DifFt As Boolean
Dim Rs As DAO.Recordset
    
    Q = FmtQQ("Select SpecNm,Ft,Lines,Tim,Sz,LdTim from Spec where SpecNm = '?'", Spnm)
    Set Rs = CurDb.OpenRecordset(Q)
    NoCur = Not FfnIsExist(Ft)
    NoLas = RsAny(Rs)
    
    Dim CurT As Date, LasT As Date 'CurTim and LasTim
    Dim CurS&, LasS&
    Dim LasFt$, LdDTim$
    CurS = FfnSz(Ft)
    CurT = FfnTim(Ft)
    If Not NoLas Then
        With Rs
            LasS = Nz(Rs!Sz, -1)
            LasT = Nz(!Tim, 0)
            LasFt = Nz(!Ft, "")
            LdDTim = DteDTim(!LdTim)
        End With
    End If
    SamTim = CurT = LasT
    CurOld = CurT < LasT
    CurNew = CurT > LasT
    SamSz = CurS = LasS
    DifSz = Not SamSz
    DifFt = Ft <> LasFt
    

Const Imported$ = "***** IMPORTED ******"
Const NoImport$ = "----- no import -----"
Const NoCur______$ = "No Ft."
Const NoLas______$ = "No Last."
Const FtDif______$ = "Ft is dif."
Const SamTimSz___$ = "Sam tim & sz."
Const SamTimDifSz$ = "Sam tim & sz. (Odd!)"
Const CurIsOld___$ = "Cur is old."
Const CurIsNew___$ = "Cur is new."
Const C$ = "|[SpecNm] [Db] [Cur-Ft] [Las-Ft] [Cur-Tim] [Las-Tim] [Cur-Sz] [Las-Sz] [Imported-Time]."

Dim Dr()
Dr = Array(Spnm, Ft, FtLines(Ft), CurT, CurS, Now)
Select Case True
Case NoCur, SamTim:
Case NoLas: DrInsRs Dr, Rs
Case DifFt, CurNew: DrUpdRs Dr, Rs
Case Else: Stop
End Select

Dim Av()
Av = Array(Spnm, DbNm(A), Ft, LasFt, CurT, LasT, CurS, LasS, LdDTim)
Select Case True
Case NoCur:            FunMsgAvLinDmp CSub, NoImport & NoCur______ & C, Av
Case NoLas:            FunMsgAvLinDmp CSub, Imported & NoLas______ & C, Av
Case DifFt:            FunMsgAvLinDmp CSub, Imported & FtDif______ & C, Av
Case SamTim And SamSz: FunMsgAvLinDmp CSub, NoImport & SamTimSz___ & C, Av
Case SamTim And DifSz: FunMsgAvLinDmp CSub, NoImport & SamTimDifSz & C, Av
Case CurOld:           FunMsgAvLinDmp CSub, NoImport & CurIsOld___ & C, Av
Case CurNew:           FunMsgAvLinDmp CSub, Imported & CurIsNew___ & C, Av
Case Else: Stop
End Select
End Sub

Function DbIsOk(A As Database) As Boolean
On Error GoTo X
DbIsOk = IsStr(A.Name)
Exit Function
X:
End Function

Sub DbKill(A As Database)
Dim F$
F = A.Name
A.Close
Kill F
End Sub

Function DbNm$(A As Database)
DbNm = ObjNm(A)
End Function

Function DbOupTny(A As Database) As String()
DbOupTny = DbqSy(A, "Select Name from MSysObjects where Name like '@*' and Type =1")
End Function

Sub DbBrw(A As Database)
Stop '
End Sub
Function DbPth$(A As Database)
DbPth = FfnPth(A.Name)
End Function

Function DbQny(A As Database) As String()
DbQny = DbqSy(A, "Select Name from MSysObjects where Type=5 and Left(Name,4)<>'MSYS' and Left(Name,4)<>'~sq_'")
End Function

Private Sub Z_DbQny()
AyDmp DbQny(CurDb)
End Sub

Function DbQryRs(A As Database, Qry) As DAO.Recordset
Set DbQryRs = A.QueryDefs(Qry).OpenRecordset
End Function

Sub DbReOpn(ODb As Database)
Dim Nm$
Nm = ODb.Name
ODb.Close
Set ODb = DAO.DBEngine.OpenDatabase(Nm)
End Sub

Sub DbResClr(A As Database, ResNm$)
A.Execute "Delete From Res where ResNm='" & ResNm & "'"
End Sub


Function DbScly(A As Database) As String()
DbScly = AySy(AyOfAy_Ay(AyMap(ItrMap(A.TableDefs, "TdScly"), "TdScly_AddPfx")))
End Function

Sub DbSetTDes(A As Database, B As TDes)
DbtDes(A, B.T) = B.Des
End Sub

Function DbSpecNy(A As DAO.Database) As String()
DbSpecNy = DbtfSy(A, "Spec", "SpecNm")
End Function

Sub DbSqyRun(A As Database, Sqy$())
Dim Q
For Each Q In AyNz(Sqy)
   A.Execute Q
Next
End Sub

Function DbSrcTny(A As Database) As String()
Dim S()
Dim T$()
T = AyQuoteSqBkt(DbTny(A))
S = AyMapPX(T, "DbtSrcTblNm", A)
DbSrcTny = AyabNonEmpBLy(T, S)
End Function

Function DbTmpTny(A As Database) As String()
DbTmpTny = AyWhPfx(DbTny(A), "#")
End Function

Function DbTny(A As Database) As String()
DbTny = FbTny(A.Name)
Exit Function
DbTny = DbqSy(A, "Select Name from MSysObjects where Type in (1,6) and Name not Like 'MSys*' and Name not Like 'f_*_Data'")
Dim T As TableDef, O$()
Dim X As DAO.TableDefAttributeEnum
X = DAO.TableDefAttributeEnum.dbHiddenObject Or DAO.TableDefAttributeEnum.dbSystemObject
For Each T In A.TableDefs
    Select Case True
    Case T.Attributes And X
    Case Else
        Push O, T.Name
    End Select
Next
DbTny = O
End Function

Function DbtTblFInfDryFny(A As Database, T$) As String()
Dim O$()
Push O, "Tbl"
Push O, "SeqNo"
PushAy O, FldInfDryFny
DbtTblFInfDryFny = O
End Function

Private Sub ZZ_DbDs()
Dim Ds As Ds
Dim Db As Database: Set Db = SampleDb_DutyPrepare
Set Ds = DbDs(Db, "Permit PermitD")
DsBrw Ds
End Sub

Private Sub ZZ_DbQny()
AyDmp DbQny(SampleDb_DutyPrepare)
End Sub
