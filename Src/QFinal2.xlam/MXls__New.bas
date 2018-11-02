Attribute VB_Name = "MXls__New"
Option Explicit
Function NewA1(Optional WsNm$ = "Sheet1") As Range
Set NewA1 = WsA1(NewWs(WsNm))
End Function

Function NewWb(Optional WsNm$) As Workbook
Dim O As Workbook
Set O = NewXls.Workbooks.Add
Set NewWb = WsWb(WsSetNm(WbFstWs(O), WsNm))
End Function

Function NewWs(Optional WsNm$) As Worksheet
Set NewWs = WsSetNm(WbFstWs(NewWb), WsNm)
End Function
Function NewWsA1(Optional WsNm$) As Range
Set NewWsA1 = WsA1(NewWs(WsNm))
End Function

Function NewXls() As Excel.Application
Set NewXls = New Excel.Application
End Function
