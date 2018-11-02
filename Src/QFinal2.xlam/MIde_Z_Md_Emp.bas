Attribute VB_Name = "MIde_Z_Md_Emp"
Option Explicit
Sub Z_MdIsEmp()
Dim M As CodeModule
'GoSub Tst1
GoSub Tst2
Exit Sub
Tst2:
    Set M = Md("Module2")
    Ept = True
    GoSub Tst
    Return
Tst1:
    '-----
    Dim T$, P As VBProject
        Set P = CurPj
        T = TmpNm
    '---
    Set M = PjAddMod(P, T)
    Ept = True
    GoSub Tst
    PjDltMd P, T
    Return
Tst:
    Act = MdIsEmp(M)
    C
    Return
End Sub

Function SrcIsEmp(A$()) As Boolean
Dim L
For Each L In AyNz(A)
    If Not ZIsEmp(L) Then Exit Function
Next
SrcIsEmp = True
End Function

Function MdIsEmp(A As CodeModule) As Boolean
Dim J%
For J = 1 To A.CountOfLines
    If Not ZIsEmp(A.Lines(J, 1)) Then Exit Function
Next
MdIsEmp = True
End Function

Private Function ZIsEmp(A) As Boolean
If HasPfx(A, "Option ") Then Exit Function
If Trim(A) <> "" Then Exit Function
ZIsEmp = True
End Function

Sub Z_PjEmpMdNy()
Brw PjEmpMdNy(CurPj)
End Sub

Function PjEmpMdNy(A As VBProject) As String()
Dim C As VBComponent, N$
N = A.Name & "."
For Each C In A.VBComponents
    Select Case C.Type
    Case vbext_ct_ClassModule, vbext_ct_StdModule
        If MdIsEmp(C.CodeModule) Then PushI PjEmpMdNy, N & C.Name & ":" & CmpTyStr(C.Type)
    End Select
Next
End Function

Function VbeEmpMdNy(A As Vbe) As String()
Dim P As VBProject
For Each P In A.VBProjects
    PushIAy VbeEmpMdNy, PjEmpMdNy(P)
Next
End Function

Sub Z_VbeEmpMdNy()
D VbeEmpMdNy(CurVbe)
End Sub
