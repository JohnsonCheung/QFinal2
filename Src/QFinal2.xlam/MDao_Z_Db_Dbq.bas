Attribute VB_Name = "MDao_Z_Db_Dbq"
Option Explicit
Sub Z()
Z_DbqV
End Sub
Private Sub Z_DbqV()
Ept = CByte(18)
Act = SqlV("Select Y from [^YM]")
C
End Sub

Function DbqV(A As Database, Q)
DbqV = RsV(DbqRs(A, Q))
End Function

Function DbtDic(A As Database, T) As Dictionary
Dim F$, S$
DbtFstSndFldNm A, T, F, S
Set DbtDic = DbqDic(A, FmtQQ("Select [?],[?] from [?]", F, S, T))
End Function
Sub DbtFstSndFldNm(A As Database, T, OFstFldNm$, OSndFldNm$)
With A.TableDefs(T)
    OFstFldNm = .Fields(0).Name
    OSndFldNm = .Fields(1).Name
End With
End Sub

Function DbqAny(A As Database, Sql) As Boolean
DbqAny = RsAny(DbqRs(A, Sql))
End Function

Sub DbqBrw(A As Database, Sql$)
DrsBrw DbqDrs(A, Sql)
End Sub

Function DbqDr(A As Database, Q$) As Variant()
DbqDr = RsDr(A.OpenRecordset(Q))
End Function

Function DbqDrs(A As Database, Q) As Drs
Set DbqDrs = RsDrs(DbqRs(A, Q))
End Function

Function DbqDry(A As Database, Q) As Variant()
DbqDry = RsDry(DbqRs(A, Q))
End Function

Function DbqDTim$(A As Database, Sql)
DbqDTim = DteDTim(DbqV(A, Sql))
End Function

Function DbqIntAy(A As Database, Q) As Integer()
DbqIntAy = RsIntAy(DbqRs(A, Q))
End Function

Function DbqLng&(A As Database, Sql)
DbqLng = DbqV(A, Sql)
End Function

Function DbqLngAy(A As Database, Sql) As Long()
DbqLngAy = RsLngAy(A.OpenRecordset(Sql))
End Function

Function DbqRs(A As Database, Q) As DAO.Recordset
Set DbqRs = A.OpenRecordset(Q)
End Function

Sub DbqRun(A As Database, Q)
A.Execute Q
End Sub

Function DbqSy(A As Database, Q) As String()
DbqSy = RsSy(A.OpenRecordset(Q))
End Function

Function DbqTim(A As Database, Q) As Date
DbqTim = DbqV(A, Q)
End Function

Function DbqVal(A As Database, Q)
DbqVal = DbqV(A, Q)
End Function
