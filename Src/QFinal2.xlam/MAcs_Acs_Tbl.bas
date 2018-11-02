Attribute VB_Name = "MAcs_Acs_Tbl"
Option Explicit

Sub AcsTCls(A As Access.Application, T)
A.DoCmd.Close acTable, T, acSaveYes
End Sub

Sub AcsTTCls(A As Access.Application, TT)
AyDoPX CvNy(TT), "AcstCls", A
End Sub

Sub AcsClsTbl(A As Access.Application)
Dim T As AccessObject
For Each T In A.CodeData.AllTables
    A.DoCmd.Close acTable, T.Name
Next
End Sub
