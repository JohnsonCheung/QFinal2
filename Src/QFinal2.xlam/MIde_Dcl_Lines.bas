Attribute VB_Name = "MIde_Dcl_Lines"
Option Explicit

Function SrcDclLinCnt%(A$())
Dim I&
    I = SrcFstMthIx(A)
    If I = -1 Then
        I = UB(A) + 1
    Else
        I = SrcMthIxTopRmkFm(A, I)
    End If
Dim O&
    For I = I - 1 To 0 Step -1
         If IsCdLin(A(I)) Then O = I + 1: GoTo X
    Next
    O = 0
X:
SrcDclLinCnt = O
End Function

Private Sub ZZ_DclTyLines()
Debug.Print DclTyLines(MdDclLy(CurMd), "AA")
End Sub
Function SrcDclLines$(A$())
SrcDclLines = JnCrLf(SrcDclLy(A))
End Function

Function SrcDclLy(A$()) As String()
If Sz(A) = 0 Then Exit Function
Dim N&
   N = SrcDclLinCnt(A)
If N = 0 Then Exit Function
SrcDclLy = AyFstNEle(A, N)
End Function

Function MdDclLinCnt%(A As CodeModule)
If A.CountOfDeclarationLines = 0 Then Exit Function
MdDclLinCnt = SrcDclLinCnt(MdSrc(A))
End Function

Function MdDclLines$(A As CodeModule)
Dim Cnt%
Cnt = MdDclLinCnt(A)
If Cnt = 0 Then Exit Function
MdDclLines = A.Lines(1, Cnt)
End Function

Function MdDclLy(A As CodeModule) As String()
MdDclLy = SplitCrLf(MdDclLines(A))
End Function
