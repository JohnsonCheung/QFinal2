Attribute VB_Name = "MDao_CurDb"
Option Explicit
Function Tny() As String()
Tny = CurDbTny
End Function
Function CurAcs() As Access.Application
Set CurAcs = Access.Application
End Function
Sub ClsCurDb()
On Error Resume Next
CurAcs.CloseCurrentDatabase
End Sub
Function FbCurDb(A) As Database
ClsCurDb
CurAcs.OpenCurrentDatabase A
Set FbCurDb = CurDb
End Function
Function CurDb() As Database
Set CurDb = CurAcs.CurrentDb
End Function

Function CurDbChk() As String()
CurDbChk = DbChk(CurDb)
End Function

Function CurDbCnSy() As String()
CurDbCnSy = DbCnSy(CurDb)
End Function

Sub CurDbDrpLnkTbl()
DbDrpLnkTbl CurDb
End Sub

Function CurDbLnkTny() As String()
CurDbLnkTny = DbLnkTny(CurDb)
End Function

Function CurDbPth$()
CurDbPth = FfnPth(CurFb)
End Function

Function CurDbScly() As String()
CurDbScly = DbScly(CurDb)
End Function

Function CurDbStru$()
CurDbStru = DbStru(CurDb)
End Function

Function CurDbTny() As String()
CurDbTny = DbTny(CurDb)
End Function
Sub RfhTmpTbl()
TblDrp "Tmp"
DbAddTmpTbl CurDb
End Sub


Function HasTbl(T$) As Boolean
HasTbl = DbHasTbl(CurDb, T)
End Function


Function DftDb(A As Database) As Database
If IsNothing(A) Then
   Set DftDb = CurDb
Else
   Set DftDb = A
End If
End Function

Function CnSy() As String()
CnSy = CurDbCnSy
End Function
