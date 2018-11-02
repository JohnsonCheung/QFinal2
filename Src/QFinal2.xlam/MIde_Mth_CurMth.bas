Attribute VB_Name = "MIde_Mth_CurMth"
Option Explicit
Function CurMth() As Mth
Dim M As CodeModule
    Set M = CurMd
Set CurMth = Mth(M, MdCurMthNm(M))
End Function

Function MdCurMthNm$(A As CodeModule)
Dim R1&, R2&, C1&, C2&
A.CodePane.GetSelection R1, C1, R2, C2
Dim K As vbext_ProcKind
MdCurMthNm = A.ProcOfLine(R1, K)
End Function

Function CurMthNm$()
CurMthNm = CurMth.Nm
End Function
