Attribute VB_Name = "MIde_Loc"
Option Explicit

Function SrcMthLCC(A$(), MthNm) As LCC
Dim R&, C&, Ix&
Ix = SrcMthNmIx(A, MthNm)
R = Ix + 1
C = InStr(A(Ix), MthNm)
SrcMthLCC = LCC(R, C + 1, C + Len(MthNm))
End Function


Function IsRRCCOutSidMd(A As RRCC, Md As CodeModule) As Boolean
IsRRCCOutSidMd = True
Dim R%
R = MdNLin(Md)
'If RRCCIsEmp(A) Then Exit Function
With A
   If .R1 > R Then Exit Function
   If .R2 > R Then Exit Function
   If .C1 > Len(Md.Lines(.R1, 1)) + 1 Then Exit Function
   If .C2 > Len(Md.Lines(.R2, 1)) + 1 Then Exit Function
End With
IsRRCCOutSidMd = False
End Function

Sub LocStr_Go(A)
LocGo LocStr_Loc(A)
End Sub

Function LocStr_Loc(A) As VbeLoc

End Function
Sub LocGo(A)

End Sub

Private Sub Z_MthFC()
Dim M As Mth: Set M = Mth(Md("ZZModule"), "YYA")
Dim Act() As FmCnt: Act = MthFC(M)
Ass Sz(Act) = 2
Ass Act(0).FmLno = 5
Ass Act(0).Cnt = 7
Ass Act(1).FmLno = 13
Ass Act(1).Cnt = 15
End Sub

Function LinMthLCC(A$, MthNm$, Lno%) As LCC
Const CSub$ = "LinMthLCC"
Dim M$: M = LinMthNm(A): If M = "" Then Er CSub, "Given [MthLin] is not a MthLin, where [MthNm] [Lno]", A, MthNm, Lno
If M <> MthNm Then Er CSub, "Given [MthLin] does not have [MthNm], where [Lno]", A, MthNm, Lno
Dim C1%, C2%
C1 = InStr(A, MthNm)
C2 = C1 + Len(MthNm)
LinMthLCC = LCC(Lno, C1, C2)
End Function

Function MdMthLoc(A As CodeModule, MthNm$) As VbeLoc
'MdMthLoc = SrcMthRRCC(MdSrc(A), MthNm)
End Function

Function RRCC_Str$(A As RRCC)
With A
'   RRCC_Str = FmtQQ("(RRCC : ? ? ? ??)", .R1, .R2, .C1, .C2, IIf(IsEmpRRCC(A), " *Empty", ""))
End With
End Function
