Attribute VB_Name = "MIde_Ens_Option"
Option Explicit
Const OptExp$ = "Option Explicit"
Const OptCmpDb$ = "Option Database"

Sub Z()
Z_EnsOptExp
Z_EnsPjOptExp
MIde_Ens_Option.EnsPjOptExp
End Sub
Sub EnsOptCmpDb(Optional MdNm$)
MdEns DftMdByNm(MdNm), OptCmpDb
End Sub
Private Sub Z_EnsOptExp()
EnsOptExp
End Sub
Sub EnsOptExp(Optional MdNm$)
MdEns DftMdByNm(MdNm), OptExp
End Sub
Private Sub Z_EnsPjOptExp()
EnsPjOptExp
End Sub
Sub EnsVbeOptExp()
Dim P As VBProject
For Each P In CurVbe.VBProjects
    PjEns P, OptExp
Next
End Sub

Sub EnsPjOptExp(Optional PjNm$)
PjEns DftPjByNm(PjNm), OptExp
End Sub
Sub EnsPjOptCmpDb(Optional PjNm$)
PjEns DftPjByNm(PjNm), OptCmpDb
End Sub

Private Sub PjEns(A As VBProject, XXX$)
Dim M
For Each M In PjMdAy(A)
    MdEns CvMd(M), OptExp
Next
End Sub

Private Sub MdEns(A As CodeModule, XXX$)
If HasXXX(A, XXX) Then Exit Sub
A.InsertLines 1, XXX
Debug.Print MdNm(A)
End Sub

Private Function HasXXX(A As CodeModule, XXX$) As Boolean
Dim I
For Each I In AyNz(MdDclLy(A))
   If HasPfx(I, XXX) Then HasXXX = True: Exit Function
Next
End Function
