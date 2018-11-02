Attribute VB_Name = "MXls__Rfh"
Option Explicit
Sub WcRfhCnStr(A As WorkbookConnection, Fb)
If IsNothing(A.OLEDBConnection) Then Exit Sub
Dim Cn$
Const Ver$ = "0.0.1"
Select Case Ver
Case "0.0.1"
    Dim S$
    S = A.OLEDBConnection.Connection
    Cn = RplBet(S, CStr(Fb), "Data Source=", ";")
Case "0.0.2"
    Cn = FbOleCnStr(Fb)
End Select
A.OLEDBConnection.Connection = Cn
End Sub

Function FxRfh(A, Fb$) As Workbook
Set FxRfh = WbRfh(FxWb(A), Fb)
End Function

Sub WsRfh(A As Worksheet)
ItrDo A.QueryTables, "QtRfh"
ItrDo A.PivotTables, "PtRfh"
ItrDo A.ListObjects, "LoRfh"
End Sub
Sub QtRfh(A As Excel.QueryTable)
A.BackgroundQuery = False
A.Refresh
End Sub

Function WbRfh(A As Workbook, Optional Fb0 = "") As Workbook
Dim Fb$
Fb = DftStr(Fb0, CurDb.Name)
If A.Connections.Count = 0 Then FbRplWbLo Fb, A
ItrDoXP A.Connections, "WcRfh", Fb
ItrDo A.PivotCaches, "PcRfh"
ItrDo A.Sheets, "WsRfh"
WbFmtAllLo A
Set WbRfh = A
'ItrDo A.Connections, "WcDlt"
End Function

Function WbRfhCnStr(A As Workbook, Fb$) As Workbook
ItrDoXP A.Connections, "WcRfhCnStr", FbOleCnStr(Fb)
Set WbRfhCnStr = A
End Function
