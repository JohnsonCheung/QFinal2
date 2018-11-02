Attribute VB_Name = "MDao_CurDb_QQ"
Option Explicit
Sub QQ(A, ParamArray Ap())
Dim Av(): Av = Ap
CurDb.Execute FmtQQAv(A, Av)
End Sub

Function QQAny(A, ParamArray Ap()) As Boolean
Dim Av(): Av = Ap
QQAny = SqlAny(FmtQQAv(A, Av))
End Function

Function QQDTim$(A, ParamArray Ap())
Dim Av(): Av = Ap
QQDTim = DbqDTim(CurDb, FmtQQAv(A, Av))
End Function

Function QQRs(A, ParamArray Ap()) As DAO.Recordset
Dim Av(): Av = Ap
Set QQRs = DbqRs(CurDb, FmtQQAv(A, Av))
End Function

Sub QQRun(QQSql, ParamArray Ap())
Dim Av(): Av = Ap
DoCmd.RunSQL FmtQQAv(QQSql, Av)
End Sub

Function QQTim(A, ParamArray Ap()) As Date
Dim Av(): Av = Ap
QQTim = DbqTim(CurDb, FmtQQAv(A, Av))
End Function

Function QQV(A, ParamArray Ap())
Dim Av(): Av = Ap
QQV = DbqV(CurDb, FmtQQAv(A, Av))
End Function
