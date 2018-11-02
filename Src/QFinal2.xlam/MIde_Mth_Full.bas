Attribute VB_Name = "MIde_Mth_Full"
Option Explicit
Sub Z()
Z_CurPjFfnAyMthFullWb
Z_MdMthFullDrs
Z_MthFullWbFmt
Z_PjFfnMthFullDrsFmLive
End Sub
Function CurPjFfnAyMthFullDrs(Optional B As WhPjMth) As Drs
Set CurPjFfnAyMthFullDrs = PjFfnAyMthFullDrs(CurPjFfnAy, B)
End Function

Sub CurPjFfnAyMthFullWb(Optional A As WhPjMth)
WbVis PjFfnAyMthFullWb(CurPjFfnAy, A)
End Sub
Function CurPjFfnAyMthFullWs() As Worksheet
Set CurPjFfnAyMthFullWs = PjFfnAyMthFullWs(CurPjFfnAy)
End Function
Function FbMthFullDrs(A, Optional B As WhPjMth) As Drs
If False Then
    Set FbMthFullDrs = VbeMthFullDrs(FbAcs(A).Vbe, B)
    Exit Function
End If
Dim Acs As New Access.Application
Debug.Print "FbMthFullDry: "; Now; " Start get Drs "; A; "==============="
Debug.Print "FbMthFullDry: "; Now; " Start open"
Set Acs = FbAcs(A)
Debug.Print "FbMthFullDry: "; Now; " Start get Drs "
Set FbMthFullDrs = VbeMthFullDrs(Acs.Vbe, B)
Debug.Print "FbMthFullDry: "; Now; " Start quit acs "
Acs.Quit acQuitSaveNone
Debug.Print "FbMthFullDry: "; Now; " acs is quit"
Set Acs = Nothing
Debug.Print "FbMthFullDry: "; Now; " acs is nothing"
End Function
Function FxaMthFullDrs(Fxa, Optional B As WhMdMth) As Drs
Dim Pj As VBProject
Set Pj = FxaPj(Fxa)
If IsNothing(Pj) Then
    FxaOpn Fxa
    Set Pj = VbeFfnPj(CurXls.Vbe, Fxa)
    If IsNothing(Pj) Then Stop
End If
Set FxaMthFullDrs = PjMthFullDrs(Pj, B)
End Function
Function MdMthFullDrsFny() As String()
MdMthFullDrsFny = AyAdd(SslSy("PjFfn Pj MdTy Md"), SrcMthIxFullDrFny)
End Function

Function PjFfnAyMthFullDrs(PjFfnAy, Optional B As WhPjMth) As Drs
Dim PjFfn
For Each PjFfn In PjFfnAy
    'PushDrs PjFfnAyMthFullDrs, PjFfnMthFullDrs(PjFfn, B)
Next
End Function

Function PjFfnAyMthFullWb(PjFfnAy$(), Optional B As WhPjMth) As Workbook
Set PjFfnAyMthFullWb = MthFullWbFmt(WsWb(PjFfnAyMthFullWs(PjFfnAy, B)))
End Function

Function PjFfnAyMthFullWs(PjFfnAy, Optional B As WhPjMth) As Worksheet
Dim O As Drs
Set O = PjFfnAyMthFullDrs(PjFfnAy, B)
Set O = DrsAddValIdCol(O, "Nm", "VbeMth")
Set O = DrsAddValIdCol(O, "Lines", "Vbe")
'Set PjFfnAyMthFullWs = WsSetCdNmAndLoNm(DrsWs(O), "MthFull")
End Function

Sub MthDrAsg(A, OShtMdy$, OShtTy$, ONm$, OPrm$, ORet$, OLinRmk$)
AyAsg A, OShtMdy, OShtTy, ONm, OPrm, ORet, OLinRmk
End Sub


Sub PjFfnEnsMthFullCache(PjFfn)
Dim D1 As Date
Dim D2 As Date
    D1 = PjFfnPjDte(PjFfn)
    D2 = PjFfnMthFullCacheDte(PjFfn)
Select Case True
Case D1 = 0:  Stop
Case D2 = 0:
Case D1 = D2: Exit Sub
Case D2 > D1: Stop
End Select
DrsRplDbt PjFfnMthFullDrsFmLive(PjFfn), MthDb, "MthCache", FmtQQ("PjFfn='?'", PjFfn)
End Sub

Function PjFfnMthFullCacheDte(PjFfn) As Date
PjFfnMthFullCacheDte = DbqVal(MthDb, FmtQQ("Select PjDte from Mth where PjFfn='?'", PjFfn))
End Function

Function PjFfnMthFullDrs(PjFfn, Optional B As WhMdMth) As Drs
PjFfnEnsMthFullCache PjFfn
Set PjFfnMthFullDrs = PjFfnMthFullDrsFmCache(PjFfn, B)
End Function

Function PjFfnMthFullDrsFmCache(PjFfn, Optional B As WhMdMth) As Drs
Dim Sql$: Sql = FmtQQ("Select * from MthCache where PjFfn='?'", PjFfn)
Set PjFfnMthFullDrsFmCache = DbqDrs(MthDb, Sql)
End Function

Function PjFfnMthFullDrsFmLive(PjFfn) As Drs
Dim V As Vbe, A, P As VBProject, PjDte As Date
Set A = PjFfnApp(PjFfn)
Set V = A.Vbe
Set P = VbePjFfnPj(V, PjFfn)
Select Case True
Case IsFb(PjFfn):  PjDte = AcsPjDte(CvAcs(A))
Case IsFxa(PjFfn): PjDte = FileDateTime(PjFfn)
Case Else: Stop
End Select
Set PjFfnMthFullDrsFmLive = DrsAddCol(PjMthFullDrs(P), "PjDte", PjDte)

If IsFb(PjFfn) Then
    CvAcs(A).CloseCurrentDatabase
End If
End Function

Function PjMthFullDrs(A As VBProject, Optional B As WhMdMth) As Drs
Dim O As Drs
Set O = Drs(MdMthFullDrsFny, PjMthFullDry(A, B))
Set O = DrsAddValIdCol(O, "Lines", "Pj")
Set O = DrsAddValIdCol(O, "Nm", "PjMth")
Set PjMthFullDrs = O
End Function

Function PjMthFullDry(A As VBProject, Optional B As WhMdMth) As Variant()
Dim M
For Each M In AyNz(PjMdAy(A, WhMdMthMd(B)))
    PushIAy PjMthFullDry, MdMthFullDry(CvMd(M), WhMdMthMth(B))
Next
End Function

Function MdMthFullDrs(A As CodeModule, Optional B As WhMth) As Drs
Set MdMthFullDrs = Drs(MdMthFullDrsFny, MdMthFullDry(A, B))
End Function

Function MdMthFullDry(A As CodeModule, Optional B As WhMth) As Variant()
Dim P As VBProject, Ffn$, Pj$, ShtTy$, Md$, MdTy$
Set P = MdPj(A)
Ffn$ = PjFfn(P)
Pj = P.Name
MdTy = MdTyStr(A)
Md = MdNm(A)
MdMthFullDry = DryInsC4(SrcMthFullDry(MdSrc(A)), Ffn, Pj, MdTy, Md)
End Function

Function VbeMthFullDrs(A As Vbe, Optional B As WhPjMth) As Drs
Dim P
For Each P In AyNz(VbePjAy(A, WhPjMth_Nm(B)))
    'PushDrs VbeMthFullDrs, PjMthFullDrs(CvPj(P), WhPjMth_MdMth(B))
Next
End Function

Function VbeMthFullWs(A As Vbe, Optional B As WhPjMth) As Worksheet
Set VbeMthFullWs = DrsWs(VbeMthFullDrs(A, B))
End Function

Private Sub Z_CurPjFfnAyMthFullWb()
WbVis PjFfnAyMthFullWb(CurPjFfnAy, WhPjMth(MdMth:=WhMdMth(WhMd("Std"))))
End Sub

Private Sub Z_MdMthFullDrs()
DrsBrw MdMthFullDrs(CurMd)
End Sub

Private Sub Z_MthFullWbFmt()
Dim Wb As Workbook
Const Fx$ = "C:\Users\user\Desktop\Vba-Lib-1\Mth.xlsx"
MthFullWbFmt WbVis(FxWb(Fx))
Stop
End Sub

Private Sub Z_PjFfnMthFullDrsFmLive()
Dim A As Drs, A1$
A1 = CurPjFfnAy()(0)
Set A = PjFfnMthFullDrsFmLive(A1)
WsVis DrsWs(A)
End Sub
Function SrcMthFullDry(A$()) As Variant()
PushNonZSz SrcMthFullDry, SrcDclFullDr(A)
Dim Ix
For Each Ix In AyNz(SrcMthIx(A))
    PushI SrcMthFullDry, SrcMthIxFullDr(A, CLng(Ix))
Next
End Function

Function SrcDclFullDr(A$()) As Variant()
Dim Dcl$
Dcl = SrcDclLines(A): If Dcl = "" Then Exit Function
Dim Cnt%
Cnt = LinCnt(Dcl)
Const FF = "Ty Nm Cnt Lines"
Dim Vy(): Vy = Array("Dcl", "*Dcl", Cnt, Dcl)
SrcDclFullDr = VyDr(Vy, FF, SrcMthIxFullDrFny)
End Function

Function SrcMthIxFullDr(A$(), MthIx&) As Variant()
Dim L$, Lines$, TopRmk$, Lno&, Cnt%
    L = SrcContLin(A, MthIx)
    Lno = MthIx + 1
    Lines = SrcMthIxLines(A, MthIx)
    Cnt = SubStrCnt(Lines, vbCrLf) + 1
    TopRmk = SrcMthIxTopRmk(A, MthIx)
Dim Dr(): Dr = LinMthDclDr(L): If Sz(Dr) = 0 Then Stop
SrcMthIxFullDr = AyAdd(Dr, Array(Lno, Cnt, Lines, TopRmk))
End Function

Function SrcMthIxFullDrFny() As String()
SrcMthIxFullDrFny = SslSy("Mdy Ty Nm Prm Ret LinRmk Lno Cnt Lines TopRmk")
End Function

Function MthFullWbFmt(A As Workbook) As Workbook
Dim Ws As Worksheet, Lo As ListObject
Set Ws = WbCdNmWs(A, "MthLoc"): If IsNothing(Ws) Then Stop
Set Lo = WsLo(Ws, "T_MthLoc"): If IsNothing(Lo) Then Stop
Dim Ws1 As Worksheet:  GoSub X_Ws1
Dim Pt1 As PivotTable: GoSub X_Pt1
Dim Lo1 As ListObject: GoSub X_Lo1
Dim Pt2 As PivotTable: GoSub X_Pt2
Dim Lo2 As ListObject: GoSub X_Lo2
Ws1.Outline.ShowLevels , 1
Set MthFullWbFmt = WsWb(Ws)
Exit Function
X_Ws1:
    Set Ws1 = WbAddWs(WsWb(Ws))
    Ws1.Outline.SummaryColumn = xlSummaryOnLeft
    Ws1.Outline.SummaryRow = xlSummaryBelow
    Return
X_Pt1:
    Set Pt1 = LoPt(Lo, WsA1(Ws1), "MdTy Nm VbeLinesId Lines", "Pj")
    PtSetRowssOutLin Pt1, "Lines"
    PtSetRowssColWdt Pt1, "VbeLinesId", 12
    PtSetRowssColWdt Pt1, "Nm", 30
    PtSetRowssRepeatLbl Pt1, "MdTy Nm"
    Return
X_Lo1:
    Set Lo1 = PtCpyToLo(Pt1, Ws1.Range("G1"))
    LoSetNm Lo1, "T_MthLines"
    LcSetWdt Lo1, "Nm", 30
    LcSetWdt Lo1, "Lines", 100
    LcSetLvl Lo1, "Lines"
    
    Return
X_Pt2:
    Set Pt2 = LoPt(Lo1, Ws1.Range("M1"), "MdTy Nm", "Lines")
    PtSetRowssRepeatLbl Pt2, "MdTy"
    Return
X_Lo2:
    Set Lo2 = PtCpyToLo(Pt2, Ws1.Range("Q1"))
    LoSetNm Lo2, "T_UsrEdtMthLoc"
    Return
Set MthFullWbFmt = A
End Function

Function LinMthDclDr(A) As Variant()
Dim L$, Mdy$, Ty$, Nm$, Prm$, Ret$, TopRmk$, LinRmk$
L = A
Mdy = ShfShtMdy(L)
Ty = ShfMthTy(L): If Ty = "" Then Exit Function
Ty = MthShtTy(Ty)
Nm = ShfNm(L)
Ret = ShfMthSfx(L)
Prm = ShfBktStr(L)
If ShfPfxSpc(L, "As") = "As" Then
    If Ret <> "" Then Stop
    Ret = ShfTerm(L)
End If
If ShfPfx(L, "'") = "'" Then
    LinRmk = L
End If
LinMthDclDr = Array(Mdy, Ty, Nm, Prm, Ret, LinRmk)
End Function

Function LinMthDrWP(A) As Variant()
Dim Dr()
Dr = LinMthDclDr(A)
If Sz(Dr) = 0 Then Exit Function
Dr(3) = AyAddCommaSpcSfxExlLas(AyTrim(SplitComma(CStr(Dr(3)))))
LinMthDrWP = Dr
End Function
