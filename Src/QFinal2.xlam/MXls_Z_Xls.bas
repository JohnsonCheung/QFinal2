Attribute VB_Name = "MXls_Z_Xls"
Option Explicit
Function Xls(Optional Vis As Boolean) As Excel.Application
Static X As Boolean, Y As Excel.Application
Dim J%
Beg:
    J = J + 1
    If J > 10 Then Stop
If Not X Then
    X = True
    Set Y = New Excel.Application
End If
On Error GoTo XX
Dim A$
A = Y.Name
Set Xls = Y
If Vis Then XlsVis Y
Exit Function
XX:
    X = True
    GoTo Beg
End Function

Function XlsAddIn(A As Excel.Application, FxaNm) As Excel.AddIn
Dim I As Excel.AddIn
For Each I In A.AddIns
    If StrIsEq(I.Name, FxaNm & ".xlam") Then Set XlsAddIn = I
Next
End Function
Function CurXls() As Excel.Application
Set CurXls = Excel.Application
End Function

Function XlsHasAddInFn(A As Excel.Application, AddInFn) As Boolean
Dim I As Excel.AddIn
Dim N$: N = UCase(AddInFn)
For Each I In A.AddIns
    If UCase(I.Name) = N Then XlsHasAddInFn = True: Exit Function
Next
End Function

Sub XlsQuit(A As Excel.Application)
ItrDo A.Workbooks, "WbClsNoSav"
A.Quit
Set A = Nothing
End Sub

Sub XlsVis(A As Excel.Application)
If Not A.Visible Then A.Visible = True
End Sub
