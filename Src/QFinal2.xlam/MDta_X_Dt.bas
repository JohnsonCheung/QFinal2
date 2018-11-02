Attribute VB_Name = "MDta_X_Dt"
Option Explicit

Sub DtBrw(A As Dt, Optional Fnn$)
AyBrw DtFmt(A), Dft(Fnn, A.DtNm)
End Sub

Function DtCsvLy(A As Dt) As String()
Dim O$()
Dim QQStr$
Dim Dr
Push O, JnComma(AyQuoteDbl(A.Fny))
For Each Dr In A.Dry
   Push O, FmtQQAv(QQStr, Dr)
Next
End Function

Function DtDrpCol(A As Dt, ColNy0, Optional DtNm$) As Dt
Dim A1 As Drs: Set A1 = DrsDrpCol(DtDrs(A), ColNy0)
Set DtDrpCol = Dt(Dft(DtNm, A.DtNm), A1.Fny, A1.Dry)
End Function

Function DtDrs(A As Dt) As Drs
Set DtDrs = Drs(A.Fny, A.Dry)
End Function

Sub DtDmp(A As Dt)
AyDmp DtFmt(A)
End Sub
Function EmpDtAy() As Dt()
End Function

Function DtIsEmp(A As Dt) As Boolean
DtIsEmp = Sz(A.Dry) = 0
End Function

Function DtReOrd(A As Dt, ColLvs$) As Dt
Dim ReOrdFny$(): ReOrdFny = SslSy(ColLvs)
Dim IxAy&(): IxAy = AyIxAy(A.Fny, ReOrdFny)
Dim OFny$(): OFny = AyReOrd(A.Fny, IxAy)
Dim ODry(): ODry = DryReOrd(A.Dry, IxAy)
DtReOrd.DtNm = A.DtNm
Set DtReOrd = Drs(OFny, ODry)
End Function
Function Dt(DtNm, Fny0, Dry()) As Dt
Dim O As New Dt
Set Dt = O.Init(DtNm, Fny0, Dry)
End Function

Function CvDt(A) As Dt
Set CvDt = A
End Function
