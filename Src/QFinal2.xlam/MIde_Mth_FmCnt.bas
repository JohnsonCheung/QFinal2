Attribute VB_Name = "MIde_Mth_FmCnt"
Option Explicit
Sub Z()
Z_MthFmCntAyWithTopRmk
End Sub
Private Sub Z_MthFmCntAyWithTopRmk()
Dim A As Mth, Ept() As FmCnt, Act() As FmCnt

Set A = DDNmMth("IdeMthFmCnt.Z_MthFmCntAyWithTopRmk")
PushObj Ept, FmCnt(2, 11)
GoSub Tst

Exit Sub
Tst:
    Act = MthFmCntAyWithTopRmk(A)
    If Not FmCntAyIsEq(Act, Ept) Then Stop
    Return
End Sub

Function MthFmCntAyWithTopRmk(A As Mth) As FmCnt()
MthFmCntAyWithTopRmk = SrcMthFmCntAyWithTopRmk(MdSrc(A.Md), A.Nm)
End Function

Function SrcMthFmCntAyWithTopRmk(A$(), MthNm$) As FmCnt()
Dim FmIx&, ToIx&, IFm, Fm&
For Each IFm In AyNz(SrcMthNmIxAy(A, MthNm))
    Fm = IFm
    FmIx = SrcMthIxTopRmkFm(A, Fm)
    ToIx = SrcMthIxTo(A, Fm)
    PushObj SrcMthFmCntAyWithTopRmk, FmCnt(FmIx + 1, ToIx - FmIx + 1)
Next
End Function
