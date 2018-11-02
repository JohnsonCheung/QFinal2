Attribute VB_Name = "MAdo_Cn_Cnq"
Option Explicit
Sub Z()
Z_CnqDrs
End Sub

Sub CnSqyRun(A As ADODB.Connection, Sqy$())
Dim Q
For Each Q In AyNz(Sqy)
   A.Execute Q
Next
End Sub

Private Sub Z_CnqDrs()
Dim Cn As ADODB.Connection: Set Cn = FxCn(SampleFx_KE24)
Dim Q$: Q = "Select * from [Sheet1$]"
Dim Drs As Drs: Drs = CnqDrs(Cn, Q)
DrsBrw Drs
End Sub

Function CnqRs(A As ADODB.Connection, Q) As ADODB.Recordset
Set CnqRs = A.Execute(Q)
End Function

Function CnqDrs(A As ADODB.Connection, Q) As Drs
Set CnqDrs = ARsDrs(CnqRs(A, Q))
End Function
