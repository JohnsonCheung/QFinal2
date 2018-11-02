Attribute VB_Name = "MIde_Mth_Op_Mov"
Option Explicit
Sub MovMth(MthPatn$, ToMdNm$)
CurMdMovMth MthPatn, Md(ToMdNm)
End Sub
Sub CurMthMov(ToMd$)
MthMov CurMth, Md(ToMd)
End Sub
Sub MdMthMov(A As CodeModule, M$, ToMd As CodeModule)
If MdHasMth(ToMd, M) Then Er "MdMthMov", "[Mth] already exist into [ToMd], where [FmMd]", M, MdNm(ToMd), MdNm(A)
MdLinesApp ToMd, MdMthLines(A, M)
MdMthRmv A, M
End Sub

Sub Z_MdMthMov()
'MdMthMov Md("Mth_"), "XX", Md("A_")
End Sub

Sub MthMov(A As Mth, ToMd As CodeModule)
Const Trc As Boolean = False
If Trc Then
Debug.Print "MthMov: Start........................\\"
Debug.Print "MthMov: From :" & MthDNm(A)
Debug.Print "MthMov: To   :" & MdDNm(ToMd) & "." & A.Nm
End If
If MthCpy(A, ToMd) Then
    Debug.Print "MthMov: Fail to copy"
    Exit Sub
End If
MthRmv A
If Trc Then
Debug.Print "MthMov: Done.......................//"
End If
End Sub

Sub Z_MthMov()
'MthMov Md("Mth_"), "XX", Md("A_")
End Sub
