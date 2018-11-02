Attribute VB_Name = "MXls_Z_Wc"
Option Explicit
Function WcAddWs(A As WorkbookConnection) As Worksheet
Dim Wb As Workbook, Ws As Worksheet, Lo As ListObject, Qt As QueryTable
Set Wb = A.Parent
Set Ws = WbAddWs(Wb, A.Name)
Ws.Name = A.Name
WcAt A, WsA1(Ws)
Set WcAddWs = Ws
End Function

Sub WcAt(A As WorkbookConnection, At As Range)
Dim Lo As ListObject
Set Lo = RgWs(At).ListObjects.Add(SourceType:=0, Source:=A.OLEDBConnection.Connection, Destination:=At)
With Lo.QueryTable
    .CommandType = xlCmdTable
    .CommandText = A.Name
    .RowNumbers = False
    .FillAdjacentFormulas = False
    .PreserveFormatting = True
    .RefreshOnFileOpen = False
    .BackgroundQuery = True
    .RefreshStyle = xlInsertDeleteCells
    .SavePassword = False
    .SaveData = True
    .AdjustColumnWidth = True
    .RefreshPeriod = 0
    .PreserveColumnInfo = True
    .ListObject.DisplayName = TblNm_LoNm(A.Name)
    .Refresh BackgroundQuery:=False
End With
End Sub

Sub WcDlt(A As WorkbookConnection)
A.Delete
End Sub

Sub WcRfh(A As WorkbookConnection, Fb)
If IsNothing(A.OLEDBConnection) Then Exit Sub
WcRfhCnStr A, Fb
A.OLEDBConnection.BackgroundQuery = False
A.OLEDBConnection.Refresh
End Sub
