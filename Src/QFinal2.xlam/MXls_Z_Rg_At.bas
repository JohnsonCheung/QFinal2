Attribute VB_Name = "MXls_Z_Rg_At"
Option Explicit
Function CellVBar(A As Range) As Range
If IsEmpty(A.Value) Then Stop
If IsEmpty(RgRC(A, 2, 1).Value) Then
    Set CellVBar = RgRC(A, 1, 1)
    Exit Function
End If
Set CellVBar = RgCRR(A, 1, 1, A.End(xlDown).Row - A.Row + 1)
End Function
