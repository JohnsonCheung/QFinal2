Attribute VB_Name = "MIde_A_Gen_QFinal2"
Option Explicit
Const FinalFxa$ = "C:\Users\User\Desktop\QFinalSln\GenQFinal2\QFinal2.xlam"

Sub GenQFinal2()
AssDupMd
CpyMd
PjSav TarPj
End Sub
Private Sub AssDupMd()
Dim DupNy$(): DupNy = AyWhDup(AssDupMd1)
Dim N, O$()
For Each N In AyNz(DupNy)
    PushIAy O, AssDupMd2(N)
Next
AyBrwThw O, "Following classes are dup"
End Sub
Private Function AssDupMd1() As String()
Dim Pj As VBProject, C As VBComponent
For Each Pj In CurVbe.VBProjects
    If Pj.Name = "QFinal1" Then GoTo Nxt
    For Each C In Pj.VBComponents
        If CmpIsClsOrMod(C) Then
            PushI AssDupMd1, C.Name
        End If
    Next
Nxt:
Next
End Function
Private Function AssDupMd2(MdNm) As String()
Dim Pj As VBProject, C As VBComponent
For Each Pj In CurVbe.VBProjects
    If Pj.Name = "QFinal1" Then GoTo Nxt
    For Each C In Pj.VBComponents
        If CmpIsClsOrMod(C) Then
            If C.Name = MdNm Then
                PushI AssDupMd2, Pj.Name & "." & C.Name
            End If
        End If
    Next
Nxt:
Next

End Function
Private Function TarPj() As VBProject
'Every Time create a new one
Static X As VBProject
If IsNothing(X) Then
    PthEns FfnPth(FinalFxa)
    FfnDlt FinalFxa
    Set X = FxaPj(FinalFxa)
End If
Set TarPj = X
End Function

Private Sub CpyMd()
Dim Pj As VBProject
For Each Pj In CurVbe.VBProjects
    If Pj.Name <> "QFinal1" Then
        CpyMd1 Pj
    End If
Next
End Sub

Private Sub CpyMd1(A As VBProject)
Dim C As VBComponent
For Each C In A.VBComponents
    If CmpIsClsOrMod(C) Then
        MdCpy C.CodeModule, TarPj
    End If
Next
End Sub
