Attribute VB_Name = "MDao_Z_Db_DbInf_Stru"
Option Explicit

Sub DbStruDmp(A As Database)
D DbStru(A)
End Sub

Function DbStru$(A As Database)
DbStru = DbttStru(A, DbTny(A))
End Function

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


Sub DbStruEns(A As Database, Stru$, B As StruBase)
Chk StruChk(Stru, B.F, B.E)
Dim S$
S = DbtStru(A, LinT1(Stru))
If S = "" Then
    DbStruCrt A, Stru, B
    Exit Sub
End If
If S = Stru Then Exit Sub
DbtReStru A, Stru, B.F, B.E
End Sub

Function Stru$()
Stru = CurDbStru
End Function

Function TFnyStruBaseTd(T, Fny$(), F As Drs, E As Dictionary) As DAO.TableDef
Dim FdAy() As DAO.Field2
Dim Fld
For Each Fld In Fny
    PushObj FdAy, FldFd(Fld, T, F, E)  '<===
Next
Set TFnyStruBaseTd = NewTd(T, FdAy)
End Function


Function StruFld(ParamArray Ap()) As Drs
Dim Dry(), Av(), Ele$, LikFF, LikFld, X
Av = Ap
For Each X In Av
    LinAsgTRst X, Ele, LikFF
    For Each LikFld In SslSy(LikFF)
        PushI Dry, Array(Ele, LikFld)
    Next
Next
Set StruFld = Drs("Ele FldLik", Dry)
End Function

Function StruFny(A) As String()
Dim L$, T$
T = LinT1(A)
L = Replace(A, "*", T)
L = Replace(L, "|", " ")
L = RmvT1(L)
StruFny = SslSy(L)
End Function

Function TTStru$(TT)
TTStru = DbttStru(CurDb, TT)
End Function

Sub TTStruDmp(TT)
D TTStru(TT)
End Sub

Function DbInfDtStru(A As Database) As Dt
Dim T$, TT, Dry(), Des$, NRec&, Stru$
For Each TT In DbTny(A)
    T = TT
    Des = DbtDes(A, T)
    Stru = RmvT1(DbtStru(A, T))
    NRec = DbtNRec(A, T)
    PushI Dry, Array(T, NRec, Des, Stru)
Next
Set DbInfDtStru = Dt("Tbl", "Tbl NRec Des", Dry)
End Function

