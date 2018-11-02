Attribute VB_Name = "MIde_Mth_Rmk_Cxt"
Option Explicit
Private Function CxtIx(Src$(), MthIx As FTIx) As FTIx
'Src -> X:MthFmno -> MthCxtFTNo
With MthIx
    Dim Ix%
    For Ix = .FmIx To .ToIx
        If Not LasChr(Src(Ix)) = "_" Then
            Ix = Ix + 1
            Exit For
        End If
    Next
    Set CxtIx = FTIx(Ix, .ToIx - 1)
End With
End Function

Sub MdRplCxt(A As CodeModule, Cxt$)
Dim N%: N = A.CountOfLines
MdClr A, IsSilent:=True
A.AddFromString Cxt
Debug.Print FmtQQ("MdRpl_Cxt: Md(?) of Ty(?) of Old-LinCxt(?) is replaced by New-Len(?) New-LinCnt(?).<-----------------", _
    MdDNm(A), MdTyNm(A), N, Len(Cxt), LinCnt(Cxt))
End Sub

Function MthCxtFT(A As Mth) As FTNo()
MthCxtFT = SrcMthCxtFT(MdBdyLy(A.Md), A.Nm)
End Function

Sub MthCxtFT_Rmk(A As Mth, Cxt As FTNo)
If IsRemarked(MdFTLy(A.Md, Cxt)) Then Exit Sub
Dim J%, L$
For J = Cxt.FmNo To Cxt.ToNo
    L = A.Md.Lines(J, 1)
    A.Md.ReplaceLine J, "'" & L
Next
A.Md.InsertLines Cxt.FmNo, "Stop" & " '"
End Sub

Sub MthCxtFT_UnRmk(A As Mth, Cxt As FTNo)
If Not IsRemarked(MdFTLy(A.Md, Cxt)) Then Exit Sub
Dim J%, L$
If Not HasPfx(A.Md.Lines(Cxt.FmNo, 1), "Stop '") Then Stop
For J = Cxt.FmNo + 1 To Cxt.ToNo
    L = A.Md.Lines(J, 1)
    If Left(L, 1) <> "'" Then Stop
    A.Md.ReplaceLine J, Mid(L, 2)
Next
A.Md.DeleteLines Cxt.FmNo, 1
End Sub

Private Sub MthCxtFTNoRmk(A As Mth, Cxt As FTNo)
If IsRemarked(MdFTLy(A.Md, Cxt)) Then Exit Sub
Dim J%, L$
For J = Cxt.FmNo To Cxt.ToNo
    L = A.Md.Lines(J, 1)
    A.Md.ReplaceLine J, "'" & L
Next
A.Md.InsertLines Cxt.FmNo, "Stop" & " '"
End Sub

Private Sub MthCxtFTNoUnRmk(A As Mth, Cxt As FTNo)
If Not IsRemarked(MdFTLy(A.Md, Cxt)) Then Exit Sub
Dim J%, L$
If Not HasPfx(A.Md.Lines(Cxt.FmNo, 1), "Stop '") Then Stop
For J = Cxt.FmNo + 1 To Cxt.ToNo
    L = A.Md.Lines(J, 1)
    If Left(L, 1) <> "'" Then Stop
    A.Md.ReplaceLine J, Mid(L, 2)
Next
A.Md.DeleteLines Cxt.FmNo, 1
End Sub


Function SrcMthFT_CxtFT(Src$(), Mth As FTIx) As FTIx
'Src -> X:MthFmno -> MthCxtFTNo
With Mth
    Dim Ix%
    For Ix = .FmIx To .ToIx
        If Not LasChr(Src(Ix)) = "_" Then
            Ix = Ix + 1
            Exit For
        End If
    Next
    Set SrcMthFT_CxtFT = FTIx(Ix, .ToIx - 1)
End With
End Function

Function MthLyCxt(MthLy$()) As String()
MthLyCxt = CxtIx(MthLy, FTNo(1, Sz(MthLy)))
End Function

Function SrcMthCxtFT(A$(), MthNm$) As FTNo()
Dim P() As FTIx
Dim Ix() As FTIx: Ix = SrcMthNmFT(A, MthNm)
SrcMthCxtFT = AyMapPXInto(Ix, "CxtIx", A, P)
End Function

Private Sub ZZ_MthCxtFT _
 _
()

Dim I
For Each I In MthCxtFT(CurMth)
    With CvFTNo(I)
        Debug.Print .FmNo, .ToNo
    End With
Next
End Sub

