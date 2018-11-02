Attribute VB_Name = "MDao__Dta"
Option Explicit
Function DsInsDbSqy(A As Ds, Db As Database) As String()
Dim Dt
For Each Dt In AyNz(A.DtAy)
   PushIAy DsInsDbSqy, DtInsDbSqy(CvDt(Dt), Db)
Next
End Function


Sub DsInsDb(A As Ds, Db As Database)
DbSqyRun Db, DsInsDbSqy(Db, A)
End Sub

Sub DtRplDb(A As Dt, Db As Database)
Dim T$: T = A.DtNm
Db.Execute QDlt(T)
DtInsDb A, Db
End Sub


Sub DsRplDb(A As Ds, Db As Database)
Dim T
For Each T In A.DtAy
    DtRplDb CvDt(T), Db
Next
End Sub

Sub DrsEnsDbt(A As Drs, Db As Database, T$)
If DbHasTbl(Db, T) Then Exit Sub
Db.Execute DrsCrtTblSql(A, T)
End Sub

Sub DrsEnsRplDbt(A As Drs, Db As Database, T$)
DrsEnsDbt A, Db, T
DrsRplDbt A, Db, T
End Sub

Sub DrsInsDbt(A As Drs, Db As Database, T$)
GoSub X
Dim Sqy$(): GoSub X_Sqy
DbSqyRun Db, Sqy
Exit Sub
X:
    Dim Dry, Fny$(), Sk$()
    Fny = A.Fny
    Dry = A.Dry
    Sk = DbtSk(Db, T)
    Return
X_Sqy:
    Dim Dr
    For Each Dr In AyNz(Dry)
        Push Sqy, InsSql(T, Fny, Dr)
    Next
    Return
End Sub

Sub DrsInsUpdDbt(A As Drs, Db As Database, T$)
GoSub X
Dim Ins As Drs, Upd As Drs: GoSub X_Ins_Upd
DrsInsDbt Ins, Db, T
DrsUpdDbt Upd, Db, T
Exit Sub
X:
    Dim Sk$()
    Dim SkIxAy&()
    Dim Fny$()
    Dim Dry()
    Sk = DbtSk(Db, T)
    Fny = A.Fny
    SkIxAy = AyIxAy(Fny, Sk)
    Dry = A.Dry
    Return
X_Ins_Upd:
    Dim IDry(), UDry(): GoSub X_IDry_UDry
    Set Ins = Drs(Fny, IDry)
    Set Upd = Drs(Fny, UDry)
    
    Return
X_IDry_UDry:
    Dim Dr, IsIns As Boolean, IsUpd As Boolean
    For Each Dr In Dry
        GoSub X_IsIns_IsUpd
        Select Case True
        Case IsIns: Push IDry, Dr
        Case IsUpd: Push UDry, Dr
        End Select
    Next
    Return
X_IsIns_IsUpd:
    IsIns = False
    IsUpd = False
    Dim SkVy(), Sql$, DbDr()
    SkVy = DrSel(CvAy(Dr), SkIxAy)
    Sql = QSel_FF_Fm_WhFny_Ay(Fny, T, Sk, SkVy)
    DbDr = DbqDr(Db, Sql)
    If Sz(DbDr) = 0 Then IsIns = True: Return
    If Not IsEqAy(DbDr, Dr) Then IsUpd = True: Return
    Return
End Sub

Private Sub Z_DtInsDbSqy()
'Tmp1Tbl_Ens
Stop
Dim Db As Database
Dim Dt As Dt: Dt = DbtDt(CurDb, "Tmp1")
Dim O$(): O = DtInsDbSqy(Dt, Db)
Stop
End Sub

Sub DtInsDb(A As Dt, Db As Database)
If Sz(A.Dry) = 0 Then Exit Sub
Dim Rs As DAO.Recordset, Dr
Set Rs = Db.TableDefs(A.DtNm).OpenRecordset
For Each Dr In A.Dry
    DrInsRs Dr, Rs
Next
Rs.Close
End Sub

Function DtInsDbSqy(A As Dt, Db As Database) As String()
Dim SimTyAy() As eSimTy
SimTyAy = DbtSimTyAy(Db, A.DtNm)
Dim ValTp$
   ValTp = SimTyAy_InsValTp(SimTyAy)
Dim Tp$
   Dim T$, F$
   T = A.DtNm
   F = JnComma(A.Fny)
   Tp = FmtQQ("Insert into [?] (?) values(?)", T, F, ValTp)
Dim O$()
   Dim Dr
   ReDim O(UB(A.Dry))
   Dim J%
   J = 0
   For Each Dr In A.Dry
       O(J) = FmtQQAv(Tp, Dr)
       J = J + 1
   Next
DtInsDbSqy = O
End Function
Sub DrsRplDbt(A As Drs, Db As Database, T$, Optional BExpr$)
Const CSub$ = "DrsRplDbt"
Dim Miss$(), FnyDb$(), FnyDrs$()
FnyDb = DbtFny(Db, T)
FnyDrs = A.Fny
Miss = AyMinus(FnyDb, FnyDrs)
If Sz(Miss) > 0 Then
    ErWh CSub, "Some field in Drs is missing in Db-T", "Drs-Fny Dbt-Fny Missing Db] T", FnyDrs, FnyDb, Miss, DbNm(Db), T
'    Er CSub, "Some field in Drs is missing in Db-T, Where [Drs-Fny] [Dbt-Fny] [Missing] [Db] [T]", FnyDrs, FnyDb, DbNm(Db), T
    'Er "DrsRplDbt", "[Db]-[T]-[Fny] has [Missing-Fny] according to [Drs-Fny]", DbNm(Db), T, FnyDb, Miss, FnyDrs
End If
Db.Execute QDlt(T, BExpr)
Dim Dr, Rs As DAO.Recordset
Set Rs = DbtRs(Db, T)
With Rs
    For Each Dr In DrsSel(A, DbtFny(Db, T)).Dry
        DrInsRs Dr, Rs
    Next
    .Close
End With
End Sub


Sub DrsRplDbtt(A As Drs, Db As Database, Tny0)
Dim T, Tny$()
Tny = CvNy(Tny0)
For Each T In Tny
    Db.Execute QDlt(T)
Next
DbSqyRun Db, DbDrsNormSqy(Db, A, Tny)
End Sub


Sub DrsUpdDbt(A As Drs, Db As Database, T$)
Dim Sqy$(): GoSub X
DbSqyRun Db, Sqy
Exit Sub
X:
    Dim Dr, Fny$(), Dry(), Sk$()
    Fny = A.Fny
    Sk = DbtSk(Db, T)
    Dry = A.Dry
    For Each Dr In AyNz(Dry)
        Push Sqy, UpdSqlFmt(T, Sk, Fny, Dr)
    Next
    Return
End Sub
