Attribute VB_Name = "MAcs_Acs_Cpy"
Option Explicit
Sub AcsCpyRf(A As Access.Application, Fb$)

End Sub
Sub AcsCpyFrm(A As Access.Application, Fb$)
Dim I As AccessObject
For Each I In A.CodeProject.AllForms
    A.DoCmd.CopyObject Fb, , acForm, I.Name
Next
End Sub

Sub AcsCpyMd(A As Access.Application, Fb$)
Dim I As AccessObject
For Each I In A.CodeProject.AllModules
    A.DoCmd.CopyObject Fb, , acModule, I.Name
Next
End Sub

Sub AcsCpyTbl(A As Access.Application, Fb$)
Dim I As AccessObject
For Each I In A.CodeData.AllTables
    If Not TblIsSys(I.Name) Then
        A.DoCmd.CopyObject Fb, , acTable, I.Name
    End If
Next
End Sub

Sub AcsCpy(A As Access.Application, Optional Fb0$)
Dim Fb$
If Fb0 = "" Then
    Fb = FfnNxt(A.CurrentDb.Name)
Else
    Fb = Fb0
End If
Ass PthIsExist(FfnPth(Fb))
Ass Not FfnIsExist(Fb)
FbCrt Fb
AcsCpyTbl A, Fb
AcsCpyFrm A, Fb
AcsCpyMd A, Fb
AcsCpyRf A, Fb
End Sub

Sub CurAcsCpy(Optional Fb0$)
AcsCpy CurAcs, Fb0
End Sub
