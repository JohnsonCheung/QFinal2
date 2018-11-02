Attribute VB_Name = "MIde_Mth_Dup"
Option Explicit
Sub Z()
Z_PjDupMthNy
Z_PjDupMthNyWithLinesId
Z_PjDupMth_Pj_Md_Mth_Dry
Z_PjPubMth_Pj_Md_Mth_Dry
End Sub
Function DupMthFNy_GpAy(A$()) As Variant()
Dim O(), J%, M$()
Dim L$ ' LasMthNm
L = Brk(A(0), ":").S1
Push M, A(0)
Dim B As S1S2
For J = 1 To UB(A)
    Set B = Brk(A(J), ":")
    If L <> B.S1 Then
        Push O, M
        Erase M
        L = B.S1
    End If
    Push M, A(J)
Next
If Sz(M) > 0 Then
    Push O, M
End If
DupMthFNy_GpAy = O
End Function

Function DupMthFNy_SamMthBdyFunFNy(A$(), Vbe As Vbe) As String()
Dim Gp(): Gp = DupMthFNy_GpAy(A)
Dim O$(), N, Ny
For Each Ny In Gp
    If DupMthFNyGp_IsDup(Ny) Then
        For Each N In Ny
            Push O, N
        Next
    End If
Next
DupMthFNy_SamMthBdyFunFNy = O
End Function

Sub DupMthFNy_ShwNotDupMsg(A$(), MthNm)
Select Case Sz(A)
Case 0: Debug.Print FmtQQ("DupMthFNy_ShwNotDupMsg: no such Fun(?) in CurVbe", MthNm)
Case 1
    Dim B As S1S2: Set B = Brk(A(0), ":")
    If B.S1 <> MthNm Then Stop
    Debug.Print FmtQQ("DupMthFNy_ShwNotDupMsg: Fun(?) in Md(?) does not have dup-Fun", B.S1, B.S2)
End Select
End Sub

Private Function ZCmpFmt(A, Optional OIx% = -1, Optional OSam% = -1, Optional InclSam As Boolean) As String()
'DupMthFNyGp is Variant/String()-of-MthFNm with all mth-nm is same
'MthFNm is MthNm in FNm-fmt
'          Mth is Prp/Sub/Fun in Md-or-Cls
'          FNm-fmt which is 'Nm:Pj.Md'
'DupMthFNm is 2 or more MthFNy with same MthNm
Ass DupMthFNyGp_IsVdt(A)
Dim J%, I%
Dim LinesAy$()
Dim UniqLinesAy$()
    LinesAy = AyMapSy(A, "FunFNm_MthLines")
    UniqLinesAy = AyWhDist(LinesAy)
Dim MthNm$: MthNm = Brk(A(0), ":").S1
Dim Hdr$(): Hdr = ZCmpFmt__1Hdr(OIx, MthNm, Sz(A))
Dim Sam$(): Sam = ZCmpFmt__2Sam(InclSam, OSam, A, LinesAy)
Dim Syn$(): Syn = ZCmpFmt__3Syn(UniqLinesAy, LinesAy, A)
Dim Cmp$(): Cmp = ZCmpFmt__4Cmp(UniqLinesAy, LinesAy, A)
ZCmpFmt = AyAddAp(Hdr, Sam, Syn, Cmp)
End Function

Private Function ZCmpFmt__1Hdr(OIx%, MthNm$, Cnt%) As String()
Dim O$(1)
O(0) = "================================================================"
Dim A$
    If OIx >= 0 Then A = FmtQQ("#DupMthNo(?) ", OIx): OIx = OIx + 1
O(1) = A + FmtQQ("DupMthNm(?) Cnt(?)", MthNm, Cnt)
ZCmpFmt__1Hdr = O
End Function

Private Function ZCmpFmt__2Sam(InclSam As Boolean, OSam%, DupMthFNyGp, LinesAy$()) As String()
If Not InclSam Then Exit Function
'{DupMthFNyGp} & {LinesAy} have same # of element
Dim O$()
Dim D$(): D = AyWhDup(LinesAy)
Dim J%, X$()
For J = 0 To UB(D)
    X = ZCmpFmt__2Sam1(OSam, D(J), DupMthFNyGp, LinesAy)
    PushAy O, X
Next
ZCmpFmt__2Sam = O
End Function

Private Function ZCmpFmt__2Sam1(OSam%, Lines$, DupMthFNyGp, LinesAy$()) As String()
Dim A1$()
    If OSam > 0 Then
        Push A1, FmtQQ("#Sam(?) ", OSam)
        OSam = OSam + 1
    End If
Dim A2$()
    Dim J%
    For J = 0 To UB(LinesAy)
        If LinesAy(J) = Lines Then
            Push A2, "Shw """ & DupMthFNyGp(J) & """"
        End If
    Next
Dim A3$()
    A3 = LinesBoxLy(Lines)
ZCmpFmt__2Sam1 = AyAddAp(A1, A2, A3)
End Function

Private Function ZCmpFmt__3Syn(UniqLinesAy$(), LinesAy$(), FunFNyGp) As String()
If Sz(UniqLinesAy) <= 1 Then Exit Function
Dim B$()
    Dim J%, I%
    Dim Lines
    For Each Lines In UniqLinesAy
        For I = 0 To UB(FunFNyGp)
            If Lines = LinesAy(I) Then
                Push B, FunFNyGp(I)
                Exit For
            End If
        Next
    Next
ZCmpFmt__3Syn = AyMapPXSy(B, "FmtQQ", "Sync_Fun ""?""")
End Function

Private Function ZCmpFmt__4Cmp(UniqLinesAy$(), LinesAy$(), FunFNyGp) As String()
If Sz(UniqLinesAy) <= 1 Then Exit Function
Dim L2$() ' = From L1 with each element with MdDNm added in front
    ReDim L2(UB(UniqLinesAy))
    Dim Fnd As Boolean, DNm$, J%, Lines$, I%
    For J = 0 To UB(UniqLinesAy)
        Lines = UniqLinesAy(J)
        Fnd = False
        For I = 0 To UB(LinesAy)
            If LinesAy(I) = Lines Then
                DNm = FunFNyGp(I)
                L2(J) = DNm & vbCrLf & StrDup("-", Len(DNm)) & vbCrLf & Lines
                Fnd = False
                GoTo Nxt
            End If
        Next
        Stop
Nxt:
    Next
ZCmpFmt__4Cmp = LinesAyLyPad(L2)
End Function

Function DupMthFNyGp_Dry(Ny$()) As Variant()
'Given Ny: Each Nm in Ny is FunNm:PjNm.MdNm
'          It has at least 2 ele
'          Each FunNm is same
'Return: N-Dr of Fields {Nm Mdy-1 Ty-1 Pj-1 Md-1 Mdy-2 Ty-2 Pj-2 Md-2 Src-1 Src-2 IsSam-Pj IsSam-Md IsSam-Src}
'        where N = Sz(Ny)-1
'        where each-field-(*-1)-of-Dr comes from Ny(0)
'        where each-field-(*-2)-of-Dr comes from Ny(1..)

Dim Md1$, Pj1$, Nm$
    FunFNm_BrkAsg Ny(0), Nm, Pj1, Md1
Dim Mth1 As Mth
    Set Mth1 = Mth(Md(Pj1 & "." & Md1), Nm)
Dim Src1$
    Src1 = MthLines(Mth1)
Dim Mdy1$, Ty1$
    MthBrkAsg Mth1, Mdy1, Ty1
Dim O()
    Dim J%
    For J = 1 To UB(Ny)
        Dim Pj2$, Nm2$, Md2$
            FunFNm_BrkAsg Ny(J), Nm2, Pj2, Md2: If Nm2 <> Nm Then Stop
        Dim Mth2 As Mth
            Set Mth2 = Mth(Md(Pj2 & "." & Md2), Nm)
            Dim Src2$
            Src2 = MthLines(Mth2)
        Dim Mdy2$, Ty2$
            MthBrkAsg Mth2, Mdy2, Ty2

        Push O, Array(Nm, _
                    Mdy1, Ty1, Pj1, Md1, _
                    Mdy2, Ty2, Pj2, Md2, Src1, Src2, Pj1 = Pj2, Md1 = Md2, Src1 = Src2)
    Next
DupMthFNyGp_Dry = O
End Function

Function DupMthFNyGp_IsDup(Ny) As Boolean
DupMthFNyGp_IsDup = AyIsAllEleEq(AyMap(Ny, "FunFNm_MthLines"))
End Function

Function DupMthFNyGp_IsVdt(A) As Boolean
If Not IsSy(A) Then Exit Function
If Sz(A) <= 1 Then Exit Function
Dim N$: N = Brk(A(0), ":").S1
Dim J%
For J = 1 To UB(A)
    If N <> Brk(A(J), ":").S1 Then Exit Function
Next
DupMthFNyGp_IsVdt = True
End Function

Function DupMthFNyGpAy_AllSameCnt%(A)
If Sz(A) = 0 Then Exit Function
Dim O%, Gp
For Each Gp In A
    If DupMthFNyGp_IsDup(Gp) Then O = O + 1
Next
DupMthFNyGpAy_AllSameCnt = O
End Function

Private Sub Z_PjDupMthNyWithLinesId()
D PjDupMthNyWithLinesId(CurPj)
End Sub
Function PjDupMthNyWithLinesId(A As VBProject) As String()
Dim Dic As New Dictionary, N
For Each N In AyNz(PjDupMthNy(A))
    PushI PjDupMthNyWithLinesId, N & "." & X1(A, N, Dic)
Next
End Function

Private Function X1%(Pj As VBProject, MdDotMthNm, Dic As Dictionary)
Dim Lines$, MdNm$, M As Mth, MthNm
BrkAsg MdDotMthNm, ".", MdNm, MthNm
Set M = Mth(PjMd(Pj, MdNm), MthNm)
Lines = MthLines(M, WithTopRmk:=True)
If Dic.Exists(Lines) Then X1 = Dic(Lines): Exit Function
Dim Ix%: Ix = Dic.Count
Dic.Add Lines, Ix
X1 = Ix
End Function

Private Sub Z_PjDupMthNy()
D PjDupMthNy(CurPj)
End Sub

Function PjDupMthNy(A As VBProject) As String()
Dim Dry()
Dry = PjDupMth_Pj_Md_Mth_Dry(A) ' PjNm MdNm MthNm
Dry = DrySrt(Dry, 2)
Dry = DrySelIxAp(Dry, 1, 2) ' MthNm MdNm
PjDupMthNy = DryMapJnDot(Dry)
End Function

Function MdPubMthNy(A As CodeModule) As String()
Const CSub$ = "MdPubMthNy"
If MdTy(A) <> vbext_ct_StdModule Then Er CSub, "Given [Md]-[Ty] must have type Std", MdNm(A), MdTyStr(A)
MdPubMthNy = AyWhDist(SrcMthNy(MdSrc(A), WhMth("Pub")))
End Function

Private Function PjFfnAyDupDry(A$()) As Variant()

End Function

Private Function VbePubMth_Pj_Md_Mth_Dry(A As VBProject) As Variant()
Dim Pj As VBProject, Dry(), P
For Each P In VbePjAy(A)
    Set Pj = P
    Dry = DryInsCol(PjPubMth_Pj_Md_Mth_Dry(Pj), PjNm(Pj))
    PushIAy VbePubMth_Pj_Md_Mth_Dry, Dry
Next
End Function

Private Sub Z_PjPubMth_Pj_Md_Mth_Dry()
Brw DryFmtss(PjPubMth_Pj_Md_Mth_Dry(CurPj))
End Sub

Private Function PjPubMth_Pj_Md_Mth_Dry(A As VBProject) As Variant()
Dim Md As CodeModule, M, MNm$, N, Pnm$
Pnm = PjNm(A)
For Each M In AyNz(PjModAy(A))
    Set Md = M
    MNm = MdNm(Md)
    For Each N In AyNz(MdMthNy(Md, WhMth(WhMdy:="Pub")))
        PushI PjPubMth_Pj_Md_Mth_Dry, Array(Pnm, MNm, N)
    Next
    If MNm = "IdeMthFbGenQFinal1" Then
        D MdMthNy(Md, WhMth(WhMdy:="Pub"))
    End If
Next
End Function

Private Sub Z_PjDupMth_Pj_Md_Mth_Dry()
Brw DryFmtss(DrySrt(PjDupMth_Pj_Md_Mth_Dry(CurPj), 2))
End Sub

Private Function PjDupMth_Pj_Md_Mth_Dry(A As VBProject) As Variant()
Dim Dry(): Dry = PjPubMth_Pj_Md_Mth_Dry(A)
PjDupMth_Pj_Md_Mth_Dry = DryWhColHasDup(Dry, 2)
End Function
Function MthNm_DupMthFNy(A) As String()
Stop '
'MthNm_DupMthFNy = VbeFunFNm(CurVbe, FunPatn:="^" & A & "$")
End Function


Sub FunCmp(FunNm$, Optional InclSam As Boolean)
D FunCmpFmt(FunNm, InclSam)
End Sub

Function FunCmpFmt(FunNm, Optional InclSam As Boolean) As String()
'Found all Fun with given name and compare within curVbe if it is same
'Note: Fun is any-Mdy Fun/Sub/Prp-in-Md
Dim O$()
Dim N$(): ' N = FunFNmAy(FunNm)
DupMthFNy_ShwNotDupMsg N, FunNm
If Sz(N) <= 1 Then Exit Function
FunCmpFmt = ZCmpFmt(N, InclSam:=InclSam)
End Function

Private Sub Z_FunCmp()
FunCmp "FfnDlt"
End Sub


Function VbeDupMdNy(A As Vbe) As String()
VbeDupMdNy = DryFmtss(DryWhDup(VbePjMdDry(A)))
End Function

Function VbeDupMthCmpLy(A As Vbe, B As WhPjMth, Optional InclSam As Boolean) As String()
Stop
Dim N$(): 'N = VbeDupMthFNm(A, B)
Dim Ay(): Ay = DupMthFNy_GpAy(N)
Dim O$(), J%
Push O, FmtQQ("Total ? dup function.  ? of them has mth-lines are same", Sz(Ay), DupMthFNyGpAy_AllSameCnt(Ay))
Dim Cnt%, Sam%
For J = 0 To UB(Ay)
    PushAy O, ZCmpFmt(Ay(J), Cnt, Sam, InclSam:=InclSam)
Next
VbeDupMthCmpLy = O
End Function

Function VbeDupMthDrs(A As Vbe, B As WhPjMth, Optional IsSamMthBdyOnly As Boolean, Optional IsNoSrt As Boolean) As Drs
Dim Fny$(), Dry()
Fny = SplitSsl("Nm Mdy-1 Ty-1 Pj-1 Md-1 Mdy-2 Ty-2 Pj-2 Md-2 Src-1 Src-2 IsSam-Pj IsSam-Md IsSam-Src")
Dry = VbeDupMthDryWh(A, B, IsSamMthBdyOnly:=IsSamMthBdyOnly)
Set VbeDupMthDrs = Drs(Fny, Dry)
End Function

Function VbeDupMthDry(A As Vbe) As Variant()
'Dim B(): B = VbeMthDry(A)
'Dim Ny$(): Ny = DryStrCol(B, 2)
'Dim N1$(): N1 = AyWhDup(Ny)
'    N1 = DupMthFNy_SamMthBdyFunFNy(N1, A)
'Dim GpAy()
'    GpAy = DupMthFNy_GpAy(N1)
'    If Sz(GpAy) = 0 Then Exit Function
'Dim O()
'    Dim Gp
'    For Each Gp In GpAy
'        PushAy O, DupMthFNyGp_Dry(CvSy(Gp))
'    Next
'VbeDupMthDry = O
End Function

Function VbeDupMthDryWh(A As Vbe, B As WhPjMth, Optional IsSamMthBdyOnly As Boolean) As Variant()
'Dim N$(): 'N = VbeFunFNm(A)
'Dim N1$(): ' N1 = MthNyWhDup(N)
'    If IsSamMthBdyOnly Then
'        N1 = DupMthFNy_SamMthBdyFunFNy(N1, A)
'    End If
'Dim GpAy()
'    GpAy = DupMthFNy_GpAy(N1)
'    If Sz(GpAy) = 0 Then Exit Function
'Dim O()
'    Dim Gp
'    For Each Gp In GpAy
'        PushAy O, DupMthFNyGp_Dry(CvSy(Gp))
'    Next
'VbeDupMthDryWh = O
End Function

Function MthNmCmpFmt(A, Optional InclSam As Boolean) As String()
Dim N$(): N = MthNm_DupMthFNy(A)
If Sz(N) > 1 Then
    MthNmCmpFmt = ZCmpFmt(N, InclSam:=InclSam)
End If
End Function

