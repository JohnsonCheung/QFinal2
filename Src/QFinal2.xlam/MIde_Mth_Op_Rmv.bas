Attribute VB_Name = "MIde_Mth_Op_Rmv"
Option Explicit
Sub Z()
Z_MthRmv
End Sub
Sub MdMthRmv(A As CodeModule, M)
MthRmv Mth(A, M)
End Sub

Private Sub Z_MthRmv()
Const N$ = "ZZModule"
Dim M As CodeModule
Dim M1 As Mth, M2 As Mth
GoSub Crt
Set M = Md(N)
Set M1 = Mth(M, "ZZRmv1")
Set M2 = Mth(M, "ZZRmv2")
MthRmv M1
MthRmv M2
MdEndTrim M
If M.CountOfLines <> 0 Then MsgBox M.CountOfLines
MdDlt M
Exit Sub
Crt:
    CurPjDltMd N
    Set M = CurPjEnsMod(N)
    MdLinesApp M, RplVBar("Property Get ZZRmv1()||End Property||Function ZZRmv2()|End Function||'|Property Let ZZRmv1(V)|End Property")
    Return
End Sub

Sub MthRmv(A As Mth)
Dim X() As FmCnt: X = MthFmCntAyWithTopRmk(A)
Const Trc As Boolean = False
If Trc Then
    Dim I
    Debug.Print "MthRmv: Mth[" & MthDNm(A) & "]"
    For Each I In X
        Debug.Print "MthRmv: "; FmCntStr(CvFmCnt(I))
    Next
    For Each I In X
        Debug.Print "MthRmv: Lines"
        With CvFmCnt(I)
            Debug.Print LinesTab(A.Md.Lines(.FmLno, .Cnt))
        End With
    Next
End If
MdRmvFC A.Md, X
End Sub
