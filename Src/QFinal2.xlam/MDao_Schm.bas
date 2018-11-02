Attribute VB_Name = "MDao_Schm"
Option Explicit
Sub DbSchmCrt(A As Database, Schm$)
Dim Er$(), StruAy$(), B As StruBase, Stru
SchmAsg Schm, _
    Er, StruAy$(), B
AyBrwThw Er
For Each Stru In StruAy
    DbStruCrt A, Stru, B
Next
End Sub

Sub DbStruCrt(A As Database, Stru, B As StruBase)
Dim Td As DAO.TableDef, Pk$, Sk$, Des$, FDesDic As Dictionary
StruAsg Stru, B, _
    Td, Pk, Sk, Des, FDesDic
A.TableDefs.Append Td
If Pk <> "" Then A.Execute Pk
If Sk <> "" Then A.Execute Sk
If Des <> "" Then DbtDes(A, Td.Name) = Des
DbtFDesDicSet A, Td.Name, FDesDic
End Sub

Private Sub StruAsg(Stru, B As StruBase, OTd As DAO.TableDef, OPk$, OSk$, ODes$, OFDes As Dictionary)
Dim T$: T = LinT1(Stru)
Dim Fny$(): Fny = StruFny(Stru)
Set OTd = FndTd(T, Fny, B.F, B.E)
If AyHas(Fny, T & "Id") Then OPk = CrtPkSql(T) Else OPk = ""
OSk = FndSkSql(Stru, T)
ODes = MayDicVal(B.TDes, T)
Set OFDes = FndFDes(T, Fny, B.FDes, B.TFDes)
End Sub

Private Function FndTd(T, Fny$(), F As Drs, E As Dictionary) As DAO.TableDef
Dim FdAy() As DAO.Field2
Dim Fld
For Each Fld In Fny
    PushObj FdAy, LookupFd(Fld, T, F, E)  '<===
Next
Set FndTd = NewTd(T, FdAy)
End Function

Private Function FndFDes(T$, Fny$(), FDes As Dictionary, TFDes As Drs) As Dictionary
Set FndFDes = New Dictionary
End Function

Private Function FndSkSql$(Stru, T)
Dim Sk$()
Sk = SslSy(RmvT1(Replace(TakBef(Stru, "|"), "*", T)))
If Sz(Sk) = 0 Then Exit Function
FndSkSql = CrtSkSql(T, Sk)
End Function

Private Sub Z_DbSchmCrt()
Dim Schm$, Db As Database
Set Db = TmpDb
Schm = _
         "Tbl A *Id *Nm | *Dte AATy Loc Expr Rmk" & _
vbCrLf & "Tbl B *Id AId *Nm | *Dte" & _
vbCrLf & "Fld Txt AATy" & _
vbCrLf & "Fld Loc Loc" & _
vbCrLf & "Fld Expr Expr" & _
vbCrLf & "Fld Mem Rmk" & _
vbCrLf & "Ele Loc Txt Rq Dft=ABC [VTxt=Loc must cannot be blank] [VRul=IsNull([Loc]) or Trim(Loc)='']" & _
vbCrLf & "Ele Expr Txt [Expr=Loc & 'abc']" & _
vbCrLf & "TDes A AA BB " & _
vbCrLf & "TDes A CC DD " & _
vbCrLf & "FDes ANm AA BB " & _
vbCrLf & "TFDes A ANm TFDes-AA-BB"
GoSub Tst
Exit Sub
Tst:
    DbSchmCrt Db, Schm
    DbBrw Db
    Stop
    Return
End Sub
Sub DbtAddFldStruBase(A As Database, Tbl$, Fld$, F As Drs, E As Dictionary)
If DbtHasFld(A, Tbl, Fld) Then Exit Sub
A.TableDefs(Tbl).Fields.Append LookupFd(Fld, Tbl, F, E)
End Sub

Sub DbtAddFnyStruBase(A As Database, Tbl$, Fny$(), F As Drs, E As Dictionary)
Dim Fld
For Each Fld In AyNz(Fny)
    DbtAddFldStruBase A, Tbl, CStr(Fld), F, E
Next
End Sub

