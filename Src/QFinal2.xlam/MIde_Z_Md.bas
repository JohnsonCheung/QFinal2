Attribute VB_Name = "MIde_Z_Md"
Option Explicit
Const CMod$ = "MIde_Z_Md."
Sub Z()
Z_MdEndTrim
Z_MdEnsPrpOnEr
Z_MdMthDDNy
Z_MdMthLinCnt
Z_MdRmvPrpOnEr
Z_MdTopRmkMthLinesAy
End Sub
Private Sub Z_MdEnmMbrCnt()
Ass MdEnmMbrCnt(Md("Ide"), "AA") = 1
End Sub

Function CvMd(A) As CodeModule
Set CvMd = A
End Function

Private Sub MdEnsZZDashAsPrv(A As CodeModule)
Dim DNm$: DNm = MdDNm(A)
If DNm = "QTool.F_Ide_MthMdy" Then
    Debug.Print FmtQQ("MdEnsZZDashAsPrv Md(?) is skipped", DNm)
    Exit Sub
End If
Dim J%, L$, By$
For J = 1 To A.CountOfLines
    L = A.Lines(J, 1)
    If LinIsPubZZDash(L) Then
        Debug.Print L
        By = MthLinEnsPrv(L)
        Debug.Print FmtQQ("MdEnsZZDashAsPrv: Md(?) Lin(?) is change to Private: [?]", DNm, J, By)
        'A.ReplaceLine J, By
    End If
Next
End Sub

Private Sub MdEnsZZDashPrvMthAsPub(A As CodeModule)
Dim DNm$: DNm = MdDNm(A)
If DNm = "QTool.F_Ide_MthMdy" Then
    Debug.Print FmtQQ("MdEnsZZDashPrvMthAsPub: Md(?) is skipped", DNm)
    Exit Sub
End If
Dim J%, L$, By$
For J = 1 To A.CountOfLines
    L = A.Lines(J, 1)
    If LinIsPrvZZDash(L) Then
        By = MthLinEnsPub(L)
        Debug.Print FmtQQ("MdEnsZZDashPrvMthAsPub: Md(?) Lin(?) is change to Public: [?]", DNm, J, By)
        A.ReplaceLine J, By
    End If
Next
End Sub

Private Sub Z_MdLinesApp()
Const MdNm$ = "Module1"
MdLinesApp CurMd, "'aa"
End Sub

Private Sub Z_MdLy()
AyBrw MdLy(CurMd)
End Sub

Private Sub Z_MdMthBdyLines()
Debug.Print Len(MdMthBdyLines(CurMd, "MdMthLines"))
Debug.Print MdMthBdyLines(CurMd, "MdMthLines")
End Sub

Private Sub Z_MdMthBdyLy()
Debug.Print Len(MdMthBdyLines(CurMd, "MdMthLines"))
Debug.Print MdMthBdyLines(CurMd, "MdMthLines")
End Sub

Private Sub Z_MdEndTrim()
Dim M As CodeModule: Set M = Md("ZZModule")
MdLinesApp M, "  "
MdLinesApp M, "  "
MdLinesApp M, "  "
MdLinesApp M, "  "
MdEndTrim M, ShwMsg:=True
Ass M.CountOfLines = 15
End Sub

Private Sub Z_MdEnsPrpOnEr()
MdEnsPrpOnEr ZZMd
End Sub

''======================================================================================
Private Sub Z_MdMthDDNy()
Dim Md1 As CodeModule
Set Md1 = Md("AAAMod")
Brw MdMthNy(Md1)
Brw MdMthDDNy(Md1)
End Sub

Private Sub Z_MdMthLinCnt()
Dim O$()
    Dim J%, M, L&, E&, A As CodeModule, Ny$()
    Set A = Md("Fct")
    Ny = MdMthNy(A)
    For Each M In Ny
        DoEvents
        L = MdMthLno(A, CStr(M))
        E = MdMthLinCnt(A, L) + L - 1
        Push O, Format(L, "0000 ") & A.Lines(L, 1)
        Push O, Format(E, "0000 ") & A.Lines(E, 1)
    Next
AyBrw O
End Sub

Private Sub Z_MdRmvPrpOnEr()
MdRmvPrpOnEr ZZMd
End Sub

Private Sub Z_MdTopRmkMthLinesAy()
Brw Jn(MdTopRmkMthLinesAy(CurMd), vbCrLf & "-----------------------------------------------" & vbCrLf)
End Sub

Private Sub ZZ_MdDrs()
'DrsBrw MdDrs(Md("IdeFeature_EnsZZ_AsPrivate"))
End Sub

Private Sub ZZ_MdMthLno()
Dim O$()
    Dim Lno, L&(), M, A As CodeModule, Ny$(), J%
    Set A = Md("Fct")
    Ny = MdMthNy(A)
    For Each M In Ny
        DoEvents
        J = J + 1
        Push L, MdMthLno(A, CStr(M))
        If J Mod 150 = 0 Then
            Debug.Print J, Sz(Ny), "Z_MdMthLno"
        End If
    Next

    For Each Lno In L
        Push O, Lno & " " & A.Lines(Lno, 1)
    Next
AyBrw O
End Sub

Private Sub ZZ_MdSrt()
Dim Md As CodeModule
GoSub X0
Exit Sub
X0:
    Dim I
'    For Each I In PjMdAy(CurPjx)
        Set Md = I
        If MdNm(Md) = "Str_" Then
            GoSub Ass
        End If
'    Next
    Return
X1:

    Return
Ass:
    Debug.Print MdNm(Md); vbTab;
    Dim BefSrt$(), AftSrt$()
    BefSrt = MdLy(Md)
    AftSrt = SplitCrLf(MdSrtLines(Md))
    If JnCrLf(BefSrt) = JnCrLf(AftSrt) Then
        Debug.Print "Is Same of before and after sorting ......"
        Return
    End If
    If Sz(AftSrt) <> 0 Then
        If AyLasEle(AftSrt) = "" Then
            Dim Pfx
            Pfx = Array("There is non-blank-line at end after sorting", "Md=[" & MdNm(Md) & "=====")
            AyBrw AyAddAp(Pfx, AftSrt)
            Stop
        End If
    End If
    Dim A$(), B$(), II
    A = AyMinus(BefSrt, AftSrt)
    B = AyMinus(AftSrt, BefSrt)
    Debug.Print
    If Sz(A) = 0 And Sz(B) = 0 Then Return
    If Sz(AyRmvEmp(A)) <> 0 Then
        Debug.Print "Sz(A)=" & Sz(A)
        AyBrw A
        Stop
    End If
    If Sz(AyRmvEmp(B)) <> 0 Then
        Debug.Print "Sz(B)=" & Sz(B)
        AyBrw B
        Stop
    End If
    Return
End Sub

Private Sub ZZ_MdSrtLines()
StrBrw MdSrtLines(CurMd)
End Sub

Private Function ZZMd() As CodeModule
Set ZZMd = CurVbe.VBProjects("StockShipRate").VBComponents("Schm").CodeModule
End Function
Function MdMthBrkAy(A As CodeModule) As Variant()
MdMthBrkAy = SrcMthBrkAy(MdSrc(A))
End Function


Function MdLisFny() As String()
MdLisFny = SplitSpc("PJ Md-Pfx Md Ty Lines NMth NMth-Pub NMth-Prv NTy NTy-Pub NTy-Prv NEnm NEnm-Pub NEnm-Prv")
End Function

Function Md(MdDNm) As CodeModule
Dim A1$(): A1 = Split(MdDNm, ".")
Select Case Sz(A1)
Case 1: Set Md = PjMd(CurPj, MdDNm)
Case 2: Set Md = PjMd(Pj(A1(0)), A1(1))
Case Else: Er "Md", "[MdDNm] should be XXX.XXX or XXX", MdDNm
End Select
End Function

Function Md_FunNy_OfPfx_ZZDash(A As CodeModule) As String()
Dim J%, O$(), L$, L1$, Is_ZFun As Boolean
For J = 1 To A.CountOfLines
    Is_ZFun = True
    L = A.Lines(J, 1)
    Select Case True
    Case HasPfx(L, "Private Sub ZZ_")
        Is_ZFun = True
        L1 = RmvPfx(L, "Sub ")
    Case HasPfx(L, "Private Sub ZZ_")
        Is_ZFun = True
        L1 = RmvPfx(L, "Sub ")
    Case Else:
        Is_ZFun = False
    End Select

    If Is_ZFun Then
        Push O, TakNm(L1)
    End If
Next
Md_FunNy_OfPfx_ZZDash = O
End Function

Function Md_TstSub_Lno%(A As CodeModule)
Dim J%
For J = 1 To A.CountOfLines
    If LinIsTstSub(A.Lines(J, 1)) Then Md_TstSub_Lno = J: Exit Function
Next
End Function

Sub MdAddFun(A As CodeModule, Nm$, Lines)
MdAddIsFun A, Nm, Lines, IsFun:=True
End Sub

Sub MdAddIsFun(A As CodeModule, Nm$, Lines, IsFun As Boolean)
Dim L$
    Dim B$
    B = IIf(IsFun, "Function", "Sub")
    L = FmtQQ("? ?()|?|End ?", B, Nm, Lines, B)
MdLinesApp A, L
MthGo Mth(A, Nm)
End Sub

Sub MdAddSub(A As CodeModule, Nm$, Lines)
MdAddIsFun A, Nm, Lines, IsFun:=False
End Sub

Sub MdAppDclLin(A As CodeModule, DclLines$)
A.InsertLines A.CountOfDeclarationLines + 1, DclLines
Debug.Print FmtQQ("MdAppDclLin: Module(?) a DclLin is inserted", MdNm(A))
End Sub

Sub MdLinesApp(A As CodeModule, Lines$)
Const CSub$ = "MdLinesApp"
If Lines = "" Then Exit Sub
Dim Bef&, Aft&, Exp&, Cnt&
Bef = A.CountOfLines
A.AddFromString Lines
Aft = A.CountOfLines
Cnt = LinesLinCnt(Lines)
Exp = Bef + Cnt
If Exp <> Aft Then
'    Er CSub, "After copy line count are inconsistents, where [Md], [LinCnt-Bef-Cpy], [LinCnt-of-lines], [Exp-LinCnt-Aft-Cpy], [Act-LinCnt-Aft-Cpy], [Lines]", _
'        MdNm(A), Bef, Cnt, Exp, Aft, Lines
End If
End Sub


Sub MdAppLy(A As CodeModule, Ly$())
MdLinesApp A, JnCrLf(Ly)
End Sub

Function MdAyWhInTy(A() As CodeModule, WhInTyAy0$) As CodeModule()
Dim TyAy() As vbext_ComponentType, Md
TyAy = CvWhCmpTy(WhInTyAy0)
Dim O() As CodeModule
For Each Md In A
    If AyHas(TyAy, CvMd(Md).Parent.Type) Then PushObj O, Md
Next
MdAyWhInTy = O
End Function

Function MdAyWhMdy(A() As CodeModule, CmpTyAy0) As CodeModule()
'MdAyWhMdy = AyWhPredXP(A, "MdIsInCmpAy", CvCmpTyAy(CmpTyAy0))
End Function

Function MdHasNoMth(A As CodeModule) As Boolean
Dim J&
For J = A.CountOfDeclarationLines + 1 To A.CountOfLines
    If LinIsMth(A.Lines(J, 1)) Then Exit Function
Next
MdHasNoMth = True
End Function

Function MdBdyLines$(A As CodeModule)
If MdHasNoMth(A) Then Exit Function
Dim L&
L = MdBdyLno(A)
MdBdyLines = A.Lines(L, A.CountOfLines)
End Function

Function MdBdyLno%(A As CodeModule)
MdBdyLno = MdDclLinCnt(A) + 1
End Function

Function MdBdyLnoCnt(A As CodeModule) As LnoCnt
Dim Lno&
Dim Cnt&
Lno = MdBdyLno(A)
Stop '
MdBdyLnoCnt = LnoCnt(Lno, Cnt)
End Function

Function MdBdyLy(A As CodeModule) As String()
MdBdyLy = SplitCrLf(MdBdyLines(A))
End Function

Function MdCanHasCd(A As CodeModule) As Boolean
Select Case MdTy(A)
Case _
    vbext_ComponentType.vbext_ct_StdModule, _
    vbext_ComponentType.vbext_ct_ClassModule, _
    vbext_ComponentType.vbext_ct_Document, _
    vbext_ComponentType.vbext_ct_MSForm
    MdCanHasCd = True
End Select
End Function


Sub MdClr(A As CodeModule, Optional IsSilent As Boolean)
With A
    If .CountOfLines = 0 Then Exit Sub
    If Not IsSilent Then Debug.Print FmtQQ("MdClr: Md(?) of lines(?) is cleared", MdNm(A), .CountOfLines)
    .DeleteLines 1, .CountOfLines
End With
End Sub

Sub MdClrBdy(A As CodeModule, Optional IsSilent As Boolean)
Stop
With A
    If .CountOfLines = 0 Then Exit Sub
    Dim N%, Lno%
        Lno = MdBdyLno(A)
        N = .CountOfLines - Lno + 1
    If N > 0 Then
        If Not IsSilent Then Debug.Print FmtQQ("MdClrBdy: Md(?) of lines(?) from Lno(?) is cleared", MdNm(A), N, Lno)
        .DeleteLines Lno, N
    End If
End With
End Sub

Sub MdClsWin(A As CodeModule)
A.CodePane.Window.Close
End Sub

Function MdCmp(A As CodeModule) As VBComponent
Set MdCmp = A.Parent
End Function

Sub MdCompare(A As CodeModule, B As CodeModule)
Dim A1 As Dictionary, B1 As Dictionary
    Set A1 = MdMthDic(A)
    Set B1 = MdMthDic(B)
DicCmpBrw A1, B1, MdDNm(A), MdDNm(B)
End Sub

Function MdCmpTy(A As CodeModule) As vbext_ComponentType
MdCmpTy = A.Parent.Type
End Function

Function MdContLin$(A As CodeModule, Lno)
Dim J&, L&
L = Lno
Dim O$: O = A.Lines(L, 1)
While LasChr(O) = "_"
    L = L + 1
    O = RmvLasChr(O) & A.Lines(L, 1)
Wend
MdContLin = O
End Function
Sub Z_MdCpy()
Dim A As CodeModule, ToPj As VBProject
'
Set ToPj = CurPj
Set A = Md("QDta.Dt")
GoSub Tst
Exit Sub
Tst:
    Dim N$
    N = MdNm(A)
    Stop
    MdCpy A, ToPj   '<====
    Ass PjHasMd(ToPj, N) = True
    PjDltMd ToPj, N
    Return
End Sub
Sub MdCpy(A As CodeModule, ToPj As VBProject, Optional ShwMsg As Boolean)
Dim N$: N = MdNm(A)
If PjHasCmp(ToPj, N) Then
    Er "MdCpy", "[Md] of [Pj] already exists in [TarPj]", N, MdPjNm(A), ToPj.Name
End If
If MdIsCls(A) Then
    MdCpy1 A, ToPj 'If ClassModule need to export and import due to the Public/Private class property can only the set by Export/Import
Else
    PjAddCmpLines ToPj, N, MdTy(A), LinesEndTrim(MdLines(A))
End If
If ShwMsg Then Debug.Print FmtQQ("MdCpy: Md(?) is copied from SrcPj(?) to TarPj(?).", MdNm(A), MdPjNm(A), ToPj.Name)
End Sub
Private Sub MdCpy1(A As CodeModule, ToPj As VBProject)
Dim T$: T = TmpFt(Fnn:=MdNm(A))
A.Parent.Export T
ToPj.VBComponents.Import T
Kill T
End Sub

Function MdIsCls(A As CodeModule) As Boolean
MdIsCls = A.Parent.Type = vbext_ct_ClassModule
End Function

Function MdTopRmkMthLinesAy(A As CodeModule) As String()
MdTopRmkMthLinesAy = SrcDicTopRmkMthLinesAy(MdMthDic(A))
End Function

Sub MdDlt(A As CodeModule)
Dim M$, P$, Pj As VBProject
    M = MdNm(A)
    Set Pj = MdPj(A)
    P = Pj.Name
Debug.Print FmtQQ("MdDlt: Before Md(?) is deleted from Pj(?)", M, P)
A.Parent.Collection.Remove A.Parent
Debug.Print FmtQQ("MdDlt: After Md(?) is deleted from Pj(?)", M, P)
End Sub

Function MdDNm$(A As CodeModule)
MdDNm = MdPjNm(A) & "." & MdNm(A)
End Function

Sub MdEndTrim(A As CodeModule, Optional ShwMsg As Boolean)
If A.CountOfLines = 0 Then Exit Sub
Dim N$: N = MdDNm(A)
Dim J%
While Trim(A.Lines(A.CountOfLines, 1)) = ""
    If ShwMsg Then FunMsgLinDmp "MdEndTrim", "[LinNo] in [Md]", A.CountOfLines, N
    A.DeleteLines A.CountOfLines, 1
    If A.CountOfLines = 0 Then Exit Sub
    If J > 1000 Then Stop
    J = J + 1
Wend
End Sub

Function MdEnmBdyLy(A As CodeModule, EnmNm$) As String()
MdEnmBdyLy = DclEnmBdyLy(MdDclLy(A), EnmNm)
End Function

Function MdEnmMbrCnt%(A As CodeModule, EnmNm$)
MdEnmMbrCnt = Sz(MdEnmMbrLy(A, EnmNm))
End Function

Function MdEnmMbrLy(A As CodeModule, EnmNm$) As String()
MdEnmMbrLy = AyWhCdLin(MdEnmBdyLy(A, EnmNm))
End Function

Function MdEnsMth(A As CodeModule, MthNm$, NewMthLines$)
Dim OldMthLines$: OldMthLines = MdMthBdyLines(A, MthNm)
If OldMthLines = NewMthLines Then
    Debug.Print FmtQQ("MdEnsMth: Mth(?) in Md(?) is same", MthNm, MdNm(A))
End If
MdMthRmv A, MthNm
MdLinesApp A, NewMthLines
Debug.Print FmtQQ("MdEnsMth: Mth(?) in Md(?) is replaced <=========", MthNm, MdNm(A))
End Function

Function MdExp(A As CodeModule)
Dim F$: F = MdSrcFfn(A)
A.Parent.Export F
Debug.Print MdNm(A)
End Function

Sub MdExport(A As CodeModule)
Dim F$: F = MdSrcFfn(A)
A.Parent.Export F
Debug.Print MdNm(A)
End Sub

Sub MdFmCntDlt(A As CodeModule, B() As FmCnt)
If Not FmCntAyIsInOrd(B) Then Stop
Dim J%
For J = UB(B) To 0 Step -1
    With B(J)
        A.DeleteLines .FmLno, .Cnt
    End With
Next
End Sub

Function MdFTLines$(A As CodeModule, X As FTNo)
Dim Cnt%: Cnt = FTNoLinCnt(X)
If Cnt = 0 Then Exit Function
MdFTLines = A.Lines(X.FmNo, Cnt)
End Function

Function MdFTLy(A As CodeModule, X As FTNo) As String()
MdFTLy = SplitCrLf(MdFTLines(A, X))
End Function

Function MdMthPfxAy(A As CodeModule) As String()
Dim N
For Each N In AyNz(MdMthNy(A))
    PushNoDup MdMthPfxAy, MthPfx(N)
Next
End Function

Function MdMthPfxCnt%(A As CodeModule)
MdMthPfxCnt = Sz(MdMthPfxAy(A))
End Function

Function MdHasMth(A As CodeModule, MthNm$) As Boolean
MdHasMth = SrcHasMth(MdBdyLy(A), MthNm)
End Function

Function MdHasNoLin(A As CodeModule) As Boolean
MdHasNoLin = A.CountOfLines = 0
End Function

Function MdHasTstSub(A As CodeModule) As Boolean
Dim I
For Each I In MdLy(A)
    If I = "Friend Sub Z()" Then MdHasTstSub = True: Exit Function
    If I = "Sub Z()" Then MdHasTstSub = True: Exit Function
Next
End Function

Function MdIsNoLin(A As CodeModule) As Boolean
MdIsNoLin = A.CountOfLines = 0
End Function

Function MdLasLin$(A As CodeModule)
Dim N%: N = MdNLin(A)
If N = 0 Then Exit Function
MdLasLin = A.Lines(N, 1)
End Function

Function MdLasLno&(A As CodeModule)
MdLasLno = A.CountOfLines
End Function

Function MdLin$(A As CodeModule, Lno&)
If Lno <= 0 Then Exit Function
With A
    If Lno <= .CountOfLines Then MdLin = .Lines(Lno, 1)
End With
End Function

Function MdLines$(A As CodeModule)
If A.CountOfLines = 0 Then Exit Function
MdLines = A.Lines(1, A.CountOfLines)
End Function

Function MdLinesByLnoCnt$(A As CodeModule, LnoCnt As LnoCnt)
With LnoCnt
    If .Cnt <= 0 Then Exit Function
    MdLinesByLnoCnt = A.Lines(.Lno, .Cnt)
End With
End Function

Function MdLno_Rmv(A As CodeModule, Lno)
If Lno = 0 Then Exit Function
MsgDmp "MdLno_Rmv: [Md]-[Lno]-[Lin] is removed", MdNm(A), Lno, A.Lines(Lno, 1)
A.DeleteLines Lno, 1
End Function

Function MdLy(A As CodeModule) As String()
MdLy = SplitCrLf(MdLines(A))
End Function

Sub MdMovMthNy(A As CodeModule, MthNy$(), ToMd As CodeModule)
If Sz(MthNy) = 0 Then Exit Sub
Dim N
For Each N In MthNy
    MdMthMov A, CStr(N), ToMd
Next
End Sub

Function MdMthBdyLines$(A As CodeModule, MthNm)
MdMthBdyLines = SrcMthBdyLines(MdBdyLy(A), MthNm)
End Function

Function MdMthBdyLy(A As CodeModule, MthNm) As String()
MdMthBdyLy = SrcMthBdyLy(MdSrc(A), MthNm)
End Function

Function MdMthLnoCntAy(A As CodeModule, MthNm$) As LnoCnt()
MdMthLnoCntAy = SrcMthLnoCntAy(MdSrc(A), MthNm)
End Function

Sub Z_MdMthLnoCntAy()
Dim A() As LnoCnt: A = MdMthLnoCntAy(Md("Md_"), "XX")
Dim J%
For J = 0 To UB(A)
    LnoCnt_Dmp A(J)
Next
End Sub

Function MdMthAy(A As CodeModule) As Mth()
Dim N
For Each N In AyNz(MdMthNy(A))
    PushObj MdMthAy, Mth(A, N)
Next
End Function



Function MdMthKeyLinesDic1(A As CodeModule) As Dictionary
'To be delete
'Dim Pfx$: Pfx = MdPjNm(A) & "." & MdNm(A) & "."
'Set MdMthKeyLinesDic = DicAddKeyPfx(SrcMthKeyLinesDic(MdSrc(A)), Pfx)
End Function

Function MdMthKy(A As CodeModule) As String()
MdMthKy = AyAddPfx(SrcMthKy(MdSrc(A)), MdDNm(A) & ".")
End Function

Function MdMthLinCnt%(A As CodeModule, MthLno&)
Dim Kd$, Lin$, EndLin$, J%
Lin = A.Lines(MthLno, 1)
Kd = LinMthKd(Lin)
If Kd = "" Then Stop
EndLin = "End " & Kd
If HasSfx(Lin, EndLin) Then
    MdMthLinCnt = 1
    Exit Function
End If
For J = MthLno + 1 To A.CountOfLines
    If HasSfx(A.Lines(J, 1), EndLin) Then
        MdMthLinCnt = J - MthLno + 1
        Exit Function
    End If
Next
Stop
End Function

Function MdMthLinDry(A As CodeModule) As Variant()
MdMthLinDry = SrcMthDclDry(MdBdyLy(A))
End Function

Function MdMthLinDryWP(A As CodeModule) As Variant()
MdMthLinDryWP = SrcMthLinDryWP(MdBdyLy(A))
End Function

Function MdMthLines$(A As CodeModule, MthNm, Optional WithTopRmk As Boolean)
MdMthLines = SrcMthLines(MdSrc(A), MthNm, WithTopRmk)
End Function

Function MdMthLno&(A As CodeModule, MthNm)
MdMthLno = 1 + SrcMthNmIx(MdSrc(A), MthNm)
End Function
Function MdMthLnoAy(A As CodeModule, MthNm) As Long()
MdMthLnoAy = AyAdd1(SrcMthNmIxAy(MdSrc(A), MthNm))
End Function

Function MdMthLnoLines$(A As CodeModule, MthLno&)
MdMthLnoLines = A.Lines(MthLno, MdMthLinCnt(A, MthLno))
End Function

Function MdMthCnt%(A As CodeModule, Optional B As WhMth)
MdMthCnt = SrcMthCnt(MdSrc(A))
End Function

Function MdMthNy(A As CodeModule, Optional B As WhMth) As String()
MdMthNy = SrcMthNy(MdBdyLy(A), B)
End Function

Function MdNEnm%(A As CodeModule)
MdNEnm = DclNEnm(MdDclLy(A))
End Function

Function MdNLin%(A As CodeModule)
MdNLin = A.CountOfLines
End Function

Function MdNm$(A As CodeModule)
MdNm = A.Parent.Name
End Function

Function MdNMth%(A As CodeModule)
MdNMth = SrcNMth(MdSrc(A))
End Function

Function MdNTy%(A As CodeModule)
MdNTy = SrcNTy(MdDclLy(A))
End Function

Function MdOptCmpDbLno%(A As CodeModule)
Dim Ay$(): Ay = MdDclLy(A)
Dim J%
For J = 0 To UB(Ay)
    If HasPfx(Ay(J), "Option Compare Database") Then MdOptCmpDbLno = J + 1: Exit Function
Next
End Function

Function MdPatnLy(A As CodeModule, Patn$) As String()
Dim Ix&(): Ix = AyWhPatnIx(MdLy(A), Patn)
Dim O$(), I, Md As CodeModule
Dim N$: N = MdNm(A)
If Sz(Ix) = 0 Then Exit Function
For Each I In Ix
   Push O, FmtQQ("MdGoLno ""?"",??' ?", N, I + 1, vbTab, A.Lines(I + 1, 1))
Next
MdPatnLy = O
End Function

Function MdPj(A As CodeModule) As VBProject
Set MdPj = A.Parent.Collection.Parent
End Function

Function MdPjNm$(A As CodeModule)
MdPjNm = MdPj(A).Name
End Function

Sub MdRmvPfx(A As CodeModule, Pfx$)
MdRen A, RmvPfx(MdNm(A), Pfx)
End Sub

Sub MdRen(A As CodeModule, NewNm$)
Const CSub$ = "MdRen"
Dim Nm$: Nm = MdNm(A)
If NewNm = Nm Then
    Debug.Print FmtQQ("MdRen: Given Md-[?] name and NewNm-[?] is same", Nm, NewNm)
    Exit Sub
End If
If PjHasMd(MdPj(A), NewNm) Then
    Debug.Print FmtQQ("MdRen: Md-[?] already exist.  Cannot rename from [?]", NewNm, MdNm(A))
    Exit Sub
End If
MdCmp(A).Name = NewNm
Debug.Print FmtQQ("MdRen: Md-[?] renamed to [?] <==========================", Nm, NewNm)
End Sub

Sub Z_MdRen()
MdRen Md("A_Rs1"), "A_Rs"
End Sub

Sub MdReportSorting(A As CodeModule)
Dim Old$: Old = MdBdyLines(A)
Dim NewLines$: NewLines = MdSrtLines(A)
Dim O$: O = IIf(Old = NewLines, "(Same)", "<====Diff")
Debug.Print MdNm(A), O
End Sub

Function MdResLy(A As CodeModule, ResNm$, Optional ResPfx$ = "ZZRes") As String()
Dim Z$()
    Z = MdMthBdyLy(A, ResPfx & ResNm)
    If Sz(Z) = 0 Then
        Er "MdResLy", "{MthNm} in {Md} is not found", ResPfx & ResNm, MdNm(A)
    End If
    Z = AyRmvFstEle(Z)
    Z = AyRmvLasEle(Z)
MdResLy = AyRmvFstChr(Z)
End Function

Function MdResStr$(A As CodeModule, ResNm$)
MdResStr = JnCrLf(MdResLy(A, ResNm))
End Function

Sub MdRmv(A As CodeModule)
Dim C As VBComponent: Set C = A.Parent
C.Collection.Remove C
End Sub

Sub MdRmvBdy(A As CodeModule)
MdRmvLnoCnt A, MdBdyLnoCnt(A)
End Sub

Sub MdRmvDcl(A As CodeModule)
If A.CountOfDeclarationLines = 0 Then Exit Sub
A.DeleteLines 1, A.CountOfDeclarationLines
End Sub

Sub MdRmvEndBlankLin(A As CodeModule)
Dim J%
While A.CountOfLines > 1
    J = J + 1
    If J > 10000 Then Stop
    If Trim(A.Lines(A.CountOfLines, 1)) <> "" Then Exit Sub
    A.DeleteLines A.CountOfLines, 1
Wend
End Sub

Sub MdRmvFC(A As CodeModule, B() As FmCnt)
If Not FmCntAyIsInOrd(B) Then Stop
Dim J%
For J = UB(B) To 0 Step -1
    With B(J)
        A.DeleteLines .FmLno, .Cnt
    End With
Next
End Sub

Sub MdRmvLines(A As CodeModule)
If A.CountOfLines = 0 Then Exit Sub
A.DeleteLines 1, A.CountOfLines
End Sub

Sub MdRmvLnoCnt(A As CodeModule, LnoCnt As LnoCnt)
With LnoCnt
    If .Cnt = 0 Then Exit Sub
    A.DeleteLines .Lno, .Cnt
End With
End Sub

Sub MdRmvLnoCntAy(A As CodeModule, LnoCntAy() As LnoCnt)
If Sz(LnoCntAy) = 0 Then Exit Sub
Dim J%, M&
M = LnoCntAy(0).Lno
For J = 1 To UB(LnoCntAy)
    If M > LnoCntAy(J).Lno Then Stop
    M = LnoCntAy(J).Lno
Next

For J = UB(LnoCntAy) To 0 Step -1
    MdRmvLnoCnt A, LnoCntAy(J)
Next
End Sub

Sub Z_MdRmvLnoCntAy()
Dim A() As LnoCnt
A = MdMthLnoCntAy(Md("Md_"), "XXX")
MdRmvLnoCntAy Md("Md_"), A
End Sub

Sub MdRmvNmPfx(A As CodeModule, Pfx$)
Dim Nm$: Nm = MdNm(A): If Not HasPfx(Nm, Pfx) Then Exit Sub
MdRen A, RmvPfx(MdNm(A), Pfx)
End Sub

Sub MdRmvOptCmpDb(A As CodeModule)
Dim I%: I = MdOptCmpDbLno(A)
If I = 0 Then Exit Sub
A.DeleteLines I
Debug.Print "MdRmvOptCmpDb: Option Compare Database at line " & I & " is removed"
End Sub


Sub MdRpl(A As CodeModule, NewMdLines$)
MdClr A
MdLinesApp A, NewMdLines
End Sub

Sub MdRplBdy(A As CodeModule, NewMdBdy$)
MdClrBdy A
MdLinesApp A, NewMdBdy
End Sub

Sub MdRplDclLy(A As CodeModule, DclLy$())
MdRmvDcl A
A.InsertLines 1, JnCrLf(DclLy)
End Sub

Sub MdRplLin(A As CodeModule, Lno, NewLin$)
With A
    .DeleteLines Lno
    .InsertLines Lno, NewLin
End With
End Sub

Sub MdSav(A As CodeModule)

End Sub

Function MdTyStr$(A As CodeModule)
MdTyStr = CmpTyStr(A.Parent.Type)
End Function

Sub MdShw(A As CodeModule)
A.CodePane.Show
End Sub

Function MdSrc(A As CodeModule) As String()
MdSrc = MdLy(A)
End Function

Function MdSrcExt$(A As CodeModule)
Dim O$
Select Case A.Parent.Type
Case vbext_ct_ClassModule: O = ".cls"
Case vbext_ct_Document: O = ".cls"
Case vbext_ct_StdModule: O = ".bas"
Case vbext_ct_MSForm: O = ".cls"
Case Else: Err.Raise 1, , "MdSrcExt: Unexpected MdCmpTy.  Should be [Class or Module or Document]"
End Select
MdSrcExt = O
End Function

Function MdSrcFfn$(A As CodeModule)
MdSrcFfn = PjSrcPth(MdPj(A)) & MdSrcFn(A)
End Function

Function MdSrcFn$(A As CodeModule)
MdSrcFn = MdNm(A) & MdSrcExt(A)
End Function

Sub Srt()
MdSrt CurMd
End Sub

Sub MdSrt(A As CodeModule)
Dim Nm$: Nm = MdNm(A)
Debug.Print "Sorting: "; AlignL(Nm, 30); " ";
Dim LinesN$: LinesN = MdSrtLines(A)
Dim LinesO$: LinesO = MdLines(A)
'Exit if same
    If LinesO = LinesN Then
        Debug.Print "<== Same"
        Exit Sub
    End If
Debug.Print "<-- Sorted";
'Delete
    Debug.Print FmtQQ("<--- Deleted (?) lines", A.CountOfLines);
    MdClr A, IsSilent:=True
'Add sorted lines
    A.AddFromString LinesN
    MdRmvEndBlankLin A
    Debug.Print "<----Sorted Lines added...."
End Sub



Function MdTy(A As CodeModule) As vbext_ComponentType
MdTy = A.Parent.Type
End Function

Function MdTyLno$(A As CodeModule, TyNm$)
MdTyLno = -1
End Function

Function MdTyNm$(A As CodeModule)
MdTyNm = CmpTyStr(MdCmpTy(A))
End Function

Function MdyIsSel(A$, MdyAy$()) As Boolean
If Sz(MdyAy) = 0 Then MdyIsSel = True: Exit Function
Dim Mdy
For Each Mdy In MdyAy
    If Mdy = "Public" Then
        If A = "" Then MdyIsSel = True: Exit Function
    End If
    If A = Mdy Then MdyIsSel = True: Exit Function
Next
End Function

Function MdyShtMdy(A)
Dim O$
Select Case A
Case "", "Public":
Case "Private": O = "Prv"
Case "Friend": O = "Frd"
Case Else: Stop
End Select
MdyShtMdy = O
End Function


Sub MdMov(A As CodeModule)
'Move the MdNm in SrcPj-(Lib_XX) to TarPj-(VbLib)
'Ass ZMd_NotExist_InTar
Dim SrcCmp As VBComponent
Dim TmpFil$
    TmpFil = TmpFfn(".txt")
'    Set SrcCmp = ZSrcCmp
    SrcCmp.Export TmpFil
    If SrcCmp.Type = vbext_ct_ClassModule Then
        'ZRmvFst4Lines TmpFil
    End If
Dim TarCmp As VBComponent
'    Set TarCmp = ZTarPj.VBComponents.Add(ZMdTy)
    TarCmp.CodeModule.AddFromFile TmpFil
'ZSrcPj.VBComponents.Remove SrcCmp
Kill TmpFil
End Sub

Sub MdLikMov(MdLikNm$)
Dim I
For Each I In AyNz(CurPjMdAy(WhMd(Nm:=WhNm("^" & MdLikNm))))
    MdMov CvMd(I)
Next
End Sub

Function MdAddOptExpLin(A As CodeModule) As CodeModule
A.InsertLines 1, "Option Explicit"
Set MdAddOptExpLin = A
End Function

