Attribute VB_Name = "MDao_Z_Db_Sql"
Option Explicit
Function SqlAny(A) As Boolean
SqlAny = DbqAny(CurDb, A)
End Function

Function SqlDry(A) As Variant()
SqlDry = DbqDry(CurDb, A)
End Function

Function SqlFny(A) As String()
SqlFny = RsFny(SqlRs(A))
End Function

Function SqlLng&(A)
SqlLng = DbqLng(CurDb, A)
End Function

Function SqlLngAy(A) As Long()
SqlLngAy = DbqLngAy(CurDb, A)
End Function


Function SqlRs(A) As DAO.Recordset
Set SqlRs = CurDb.OpenRecordset(A)
End Function

Sub SqlRun(A)
CurDb.Execute A
End Sub

Function SqlStrCol(A) As String()
SqlStrCol = RsStrCol(CurDb.OpenRecordset(A))
End Function

Function SqlSy(A) As String()
SqlSy = DbqSy(CurDb, A)
End Function

Function SqlV(A)
SqlV = DbqV(CurDb, A)
End Function

Private Sub ZZ_SqlFny()
Const S$ = "SELECT qSku.*" & _
" FROM [N:\SAPAccessReports\DutyPrepay5\DutyPrepay5 (With Import).accdb].[qSku] AS qSku;"
AyDmp SqlFny(S)
End Sub

Private Sub ZZ_SqlRs()
Const S$ = "SELECT qSku.*" & _
" FROM [N:\SAPAccessReports\DutyPrepay5\DutyPrepay5 (With Import).accdb].[qSku] AS qSku;"
AyBrw RsCsvLy(SqlRs(S))
End Sub

Private Sub ZZ_SqlSy()
AyDmp SqlSy("Select Distinct UOR from [>Imp]")
End Sub

