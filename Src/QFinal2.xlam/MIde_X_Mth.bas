Attribute VB_Name = "MIde_X_Mth"
Option Explicit
Sub Z()
End Sub
Function Mth(A As CodeModule, MthNm) As Mth
Set Mth = New Mth
With Mth
    Set .Md = A
    .Nm = MthNm
End With
End Function

Sub MthAyMov(A() As Mth, ToMd As CodeModule)
AyDoXP A, "MthMov", ToMd
End Sub

Function MthDNm$(A As Mth)
MthDNm = MdDNm(A.Md) & "." & A.Nm
End Function
Function MthFul$(MthNm$)
MthFul = VbeMthMdDNm(CurVbe, MthNm)
End Function

Function MthKeyDrFny() As String()
MthKeyDrFny = SslSy("PjNm MdNm Priority Nm Ty Mdy")
End Function

Function MthKy_Sq(A$()) As Variant()
Dim O(), J%
ReDim O(1 To Sz(A) + 1, 1 To 6)
SqSetRow O, 1, MthKeyDrFny
For J = 0 To UB(A)
    SqSetRow O, J + 2, Split(A(J), ":")
Next
MthKy_Sq = O
End Function

Function MthIsExist(A As Mth) As Boolean
MthIsExist = MdHasMth(A.Md, A.Nm)
End Function


Function CvMth(A) As Mth
Set CvMth = A
End Function

Function IsMthFun(A As Mth) As Boolean
IsMthFun = IsMod(A.Md)
End Function


Function MthLno&(A As Mth)
MthLno = MdMthLno(A.Md, A.Nm)
End Function

Function MthLnoAy(A As Mth) As Integer()
MthLnoAy = AyAdd1(SrcMthNmIx(MdSrc(A.Md), A.Nm))
End Function

Function MthLnoCntAy(A As CodeModule, MthNm$) As LnoCnt()
MthLnoCntAy = SrcMthLnoCntAy(MdSrc(A), MthNm)
End Function

Sub Z_MthLnoCntAy()
Dim A() As LnoCnt: A = MthLnoCntAy(Md("Md_"), "XX")
Dim J%
For J = 0 To UB(A)
    LnoCnt_Dmp A(J)
Next
End Sub

Function MthMdDNm$(A As Mth)
MthMdDNm = MdDNm(A.Md)
End Function

Function MthMdNm$(A As Mth)
MthMdNm = MdNm(A.Md)
End Function


Function MthPjNm$(A As Mth)
MthPjNm = MdPjNm(A.Md)
End Function


Sub MthRmk(A As Mth)
Dim P() As FTNo: P = MthCxtFT(A)
Dim J%
For J = UB(P) To 0 Step -1
    MthCxtFT_Rmk A, P(J)
Next
End Sub

Sub MthRpl(A As Mth, By$)
MthRmv A
MdLinesApp A.Md, By
End Sub

Function MthIsPub(A As Mth) As Boolean
Const CSub$ = "MthIsPub"
Dim L$: L = MthDcl(A): If L = "" Then Er CSub, "Given [Mth] has a blank method line", A
MthIsPub = LinIsPubMth(L)
End Function
