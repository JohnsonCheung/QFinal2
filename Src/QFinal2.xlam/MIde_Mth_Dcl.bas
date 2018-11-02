Attribute VB_Name = "MIde_Mth_Dcl"
Option Explicit
Function CurMdMthDclLy() As String()
CurMdMthDclLy = MdMthDclLy(CurMd)
End Function

Function CurPjMthDclLy() As String()
CurPjMthDclLy = PjMthDclLy(CurPj)
End Function

Function MdMthDclLy(A As CodeModule) As String()
If MdIsNoLin(A) Then Exit Function
Dim O$(), J%
For J = 1 To A.CountOfLines
    If LinIsMth(A.Lines(J, 1)) Then
        Push O, MdContLin(A, J)
    End If
Next
MdMthDclLy = O
End Function

Function PjMthDclLy(A As VBProject) As String()
Dim I, O$(), N$, M As CodeModule
For Each I In PjMdAy(A)
    Set M = I
    N = MdTyStr(M) & "." & MdNm(M) & "."
    PushAy O, AyAddPfx(MdMthDclLy(M), N)
Next
PjMthDclLy = O
End Function


Function MdMthDclAy(A As CodeModule) As String()
MdMthDclAy = SrcMthDclAy(MdSrc(A))
End Function

Function MthDcl$(A As Mth)
MthDcl = SrcMthDcl(MdBdyLy(A.Md), A.Nm)
End Function


Function SrcMthDcl$(A$(), MthNm$)
SrcMthDcl = SrcContLin(A, SrcMthNmIx(A, MthNm))
End Function
Function SrcMthDclDry(A$()) As Variant()
Dim L
For Each L In AyNz(A)
    PushNonZSz SrcMthDclDry, LinMthDclDr(L)
Next
End Function


Function SrcMthDclAy(A$(), Optional B As WhMth) As String()
Dim O$(), J&
For J = 0 To UB(A)
    If LinIsMthWh(A(J), B) Then
        Push SrcMthDclAy, SrcContLin(A, J)
    End If
Next
End Function
