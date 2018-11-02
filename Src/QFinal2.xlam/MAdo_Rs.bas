Attribute VB_Name = "MAdo_Rs"
Option Explicit
Const SampleFb_DutyPrepare$ = ""
Function ARsDrs(A As ADODB.Recordset) As Drs
Set ARsDrs = Drs(ARsFny(A), ARsDry(A))
End Function

Function ARsDry(A As ADODB.Recordset) As Variant()
While Not A.EOF
    PushI ARsDry, AFdsDr(A.Fields)
    A.MoveNext
Wend
End Function

Private Sub Z_ARsDry()
Dim S$
Const Q$ = "Select * from KE24"
S = "GRANT SELECT ON MSysObjects TO Admin;"
'CurrentProject.Connection.Execute S
DryBrw ARsDry(CnqRs(FbCn(SampleFb_DutyPrepare), Q))
End Sub

Function ARsFny(A As ADODB.Recordset) As String()
ARsFny = AFdsFny(A.Fields)
End Function

Function ARsIntAy(A As ADODB.Recordset, Optional Col = 0) As Integer()
ARsIntAy = ARsInto(A, EmpIntAy, Col)
End Function

Function ARsInto(A As ADODB.Recordset, OInto, Optional Col = 0)
ARsInto = AyCln(OInto)
With A
    While Not .EOF
        PushI ARsInto, Nz(.Fields(Col).Value, Empty)
        .MoveNext
    Wend
    .Close
End With
End Function

Function ARsSy(A As ADODB.Recordset, Optional Col = 0) As String()
ARsSy = ARsInto(A, EmpSy, Col)
End Function
