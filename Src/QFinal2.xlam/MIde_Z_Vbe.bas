Attribute VB_Name = "MIde_Z_Vbe"
Option Explicit
Sub Z()
Z_VbeAyMthWs
Z_VbeMthLinDry
Z_VbeMthLinDryWP
End Sub
Function VbeAyMthDrs(A() As Vbe) As Drs
Dim I, R%, M As Drs
For Each I In AyNz(A)
    Set M = DrsInsCol(VbeMthDrs(CvVbe(I)), "Vbe", R)
    If R = 0 Then
        Set VbeAyMthDrs = M
    Else
        Stop
        PushObj VbeAyMthDrs, M
        Stop
    End If
    R = R + 1
    Debug.Print R; "<=== VbeAyMthDrs"
Next
End Function

Function VbeAyMthWs(A() As Vbe) As Worksheet
Set VbeAyMthWs = DrsWs(VbeAyMthDrs(A))
End Function

Function VbeBarNy(A As Vbe) As String()
VbeBarNy = ItrNy(A.CommandBars)
End Function

Function VbeCmdBarAy(A As Vbe) As Office.CommandBar()
Dim O() As Office.CommandBar
Dim I
For Each I In A.CommandBars
   PushObj O, I
Next
VbeCmdBarAy = O
End Function

Function VbeCmdBarNy(A As Vbe) As String()
VbeCmdBarNy = ItrNy(A.CommandBars)
End Function

Sub VbeCompile(A As Vbe)
ItrDo A.VBProjects, "PjCompile"
End Sub

Sub VbeCrtBar(A As Vbe, Nm$)

End Sub

Sub VbeDmpIsSaved(A As Vbe)
Dim I As VBProject
For Each I In A.VBProjects
    Debug.Print I.Saved, I.BuildFileName
Next
End Sub

Sub VbeEnsZZDashPubMthAsPrivate(A As Vbe)
AyDo VbePjAy(A), "PjEnsZZDashPubMthAsPrivate"
End Sub

Sub VbeExp(A As Vbe)
OyDo VbePjAy(A), "PjExport"
End Sub

Function VbeFfnPj(A As Vbe, PjFfn) As VBProject
Stop '
End Function

Function VbeFstQPj(A As Vbe) As VBProject
Dim I
For Each I In A.VBProjects
    If FstChr(CvPj(I).Name) = "Q" Then
        Set VbeFstQPj = I
        Exit Function
    End If
Next
End Function

Function VbeHasBar(A As Vbe, Nm$) As Boolean
VbeHasBar = ItrHasNm(A.CommandBars, Nm)
End Function

Function VbeHasPj(A As Vbe, PjNm) As Boolean
VbeHasPj = ItrHasNm(A.VBProjects, PjNm)
End Function

Function VbeHasPjFfn(A As Vbe, Ffn) As Boolean
Dim P As VBProject
For Each P In A.VBProjects
    If PjFfn(P) = Ffn Then VbeHasPjFfn = True: Exit Function
Next
End Function

Function VbeMnuBar(A As Vbe) As CommandBar
Set VbeMnuBar = A.CommandBars("Menu Bar")
End Function

Function VbeMthDot(A As Vbe, Optional MthRe As RegExp, Optional MthExlAy$, Optional WhMdyAy, Optional WhMthKd0$, Optional PjRe As RegExp, Optional PjExlAy$, Optional MdRe As RegExp, Optional MdExlAy$)
Stop '
'Dim O$(), P
'For Each P In AyNz(VbePjAy(A, PjPatn, PjExlAy))
'    PushAy O, PjMthDot(CvPj(P), MthPatn, MthExlAy, MdPatn, MdExlAy, WhMdyA, WhMthKd0)
'Next
'VbeMthDot = O
End Function

Function VbeMthDrs(A As Vbe, Optional B As WhMth) As Drs
'Dim O As Drs, O1 As Drs, O2 As Drs
'Set O = Drs("Pj Md Mdy Ty Nm Lines", VbeMthDry(A))
'Set O1 = DrsAddValIdCol(O, "Nm")
'Set O2 = DrsAddValIdCol(O1, "Lines")
'Set VbeMthDrs = O2
End Function

Function VbeMthFny() As String()
VbeMthFny = ApSy("Pj", "Md", "Mdy", "Ty", "Nm", "Lines")
End Function

Function VbeMthFx$()
VbeMthFx = FfnNxt(CurPjPth & "VbeMth.xlsx")
End Function

Function VbeMthKy(A As Vbe, Optional IsWrap As Boolean) As String()
Dim O$(), I
For Each I In VbePjAy(A)
    PushAy O, PjMthKy(CvPj(I), IsWrap)
Next
VbeMthKy = O
End Function

Function VbeMthLinDry(A As Vbe) As Variant()
Dim P
For Each P In AyNz(VbePjAy(A))
    PushAy VbeMthLinDry, PjMthLinDry(CvPj(P))
Next
End Function

Function VbeMthLinDryWP(A As Vbe) As Variant()
Dim P
For Each P In AyNz(VbePjAy(A))
    PushIAy VbeMthLinDryWP, PjMthLinDryWP(CvPj(P))
Next
End Function

Function VbeMthMdDNm$(A As Vbe, MthNm$)
Dim Pj As VBProject, P, M, Md As CodeModule
For Each P In AyNz(VbePjAy(A))
    Set Pj = P
    For Each M In PjMdAy(Pj)
        Set Md = M
        If MdHasMth(Md, MthNm) Then VbeMthMdDNm = MdDNm(Md) & "." & MthNm: Exit Function
    Next
Next
End Function

Function VbeMthNy(A As Vbe, Optional B As WhPjMth) As String()
Dim I, W As WhMdMth
Set W = WhPjMth_MdMth(B)
For Each I In AyNz(VbePjAy(A, WhPjMth_Nm(B)))
    PushIAy VbeMthNy, PjMthNy(CvPj(I), W)
Next
End Function

Function VbeMthWb(A As Vbe) As Workbook
Set VbeMthWb = WbVis(WbSavAs(WsWb(VbeMthWs(A)), VbeMthFx))
End Function

Function VbeMthWs(A As Vbe) As Worksheet
Set VbeMthWs = DrsWs(VbeMthDrs(A))
End Function

Function VbePj(A As Vbe, Pj$) As VBProject
Set VbePj = A.VBProjects(Pj)
End Function

Function VbePjAy(A As Vbe, Optional B As WhNm) As VBProject()
VbePjAy = ItrWhNmInto(A.VBProjects, B, VbePjAy)
End Function

Function VbePjFfn_Pj(A As Vbe, Ffn) As VBProject
Dim I
For Each I In A.VBProjects ' Cannot use VbePjAy(A), should use A.VBProjects
                           ' due to VbePjAy(X).FileName gives error
                           ' but (Pj in A.VBProjects).FileName is OK
    Debug.Print PjFfn(CvPj(I))
    If StrIsEq(PjFfn(CvPj(I)), Ffn) Then
        Set VbePjFfn_Pj = I
        Exit Function
    End If
Next
End Function

Function VbePjFfnAy(A As Vbe) As String()
Dim P As VBProject
For Each P In A.VBProjects
    PushNonBlankStr VbePjFfnAy, PjFfn(P)
Next
End Function

Function VbePjFfnPj(A As Vbe, Ffn) As VBProject
Dim P As VBProject
For Each P In A.VBProjects
    If StrComp(PjFfn(P), Ffn, vbTextCompare) = 0 Then
        Set VbePjFfnPj = P
        Exit Function
    End If
Next
End Function

Function VbePjMdDry(A As Vbe) As Variant()
Dim O(), P, C, Pnm$, Pj As VBProject
For Each P In VbePjAy(A)
    Set Pj = P
    Pnm = PjNm(Pj)
    For Each C In PjCmpAy(Pj)
        Push O, Array(Pnm, CvCmp(C).Name)
    Next
Next
VbePjMdDry = O
End Function

Function VbePjMdFmt(A As Vbe) As String()
VbePjMdFmt = DryFmtss(VbePjMdDry(A))
End Function

Sub VbePjMdFmtBrw(A As Vbe)
Brw VbePjMdFmt(A)
End Sub

Function VbePjNy(A As Vbe, Optional B As WhNm) As String()
VbePjNy = ItrNy(VbePjAy(A, B))
End Function

Function VbePjNyWh(A As Vbe, B As WhNm) As String()
VbePjNyWh = AyWhNm(VbePjNy(A), B)
End Function

Function VbePjNyWhMd(A As Vbe, MdPatn$) As String()
Dim I, Re As New RegExp
Re.Pattern = MdPatn
For Each I In VbePjAy(A)
    If PjHasCmpWhRe(CvPj(I), Re) Then
        Push VbePjNyWhMd, CvPj(I).Name
    End If
Next
End Function

Sub VbeSav(A As Vbe)
ItrDo A.VBProjects, "PjSav"
End Sub

Function VbeSrc(A As Vbe) As String()
Dim P
For Each P In A.VBProjects
    PushAy VbeSrc, PjSrc(CvPj(P))
Next
End Function

Function VbeSrcPth(A As Vbe)
Dim Pj As VBProject:
Set Pj = VbeFstQPj(A)
Dim Ffn$: Ffn = PjFfn(Pj)
If Ffn = "" Then Exit Function
VbeSrcPth = FfnPth(Pj.Filename)
End Function

Sub VbeSrcPthBrw(A As Vbe)
PthBrw VbeSrcPth(A)
End Sub

Sub VbeSrt(A As Vbe)
Dim I
For Each I In VbePjAy(A)
    PjSrt CvPj(I)
Next
End Sub

Sub VbeSrtRptBrw(A As Vbe)
Brw VbeSrtRptFmt(A)
End Sub

Function VbeSrtRptFmt(A As Vbe) As String()
Dim Ay() As VBProject: Ay = VbePjAy(A)
Dim O$(), I, M As VBProject
For Each I In Ay
    Set M = I
    PushAy O, PjSrtRpt(M)
Next
VbeSrtRptFmt = O
End Function

Function VbeVisWinCnt%(A As Vbe)
VbeVisWinCnt = ItrCntTruePrp(A.Windows, "Visible")
End Function

Private Sub Z_VbeAyMthWs()
WsVis VbeAyMthWs(ZZVbeAy)
End Sub

Private Sub Z_VbeMthLinDry()
Brw DryFmtss(VbeMthLinDry(CurVbe))
End Sub

Private Sub Z_VbeMthLinDryWP()
Brw DryFmtssWrp(VbeMthLinDryWP(CurVbe))
End Sub

Private Sub ZZ_VbeDmpIsSaved()
VbeDmpIsSaved CurVbe
End Sub

Private Sub ZZ_VbeDupMthCmpLy()
Brw VbeDupMthCmpLy(CurVbe, WhEmpPjMth)
End Sub

Private Sub ZZ_VbeFunPfx()
'D VbeMthPfx(CurVbe)
End Sub

Private Sub ZZ_VbeMthNy()
Brw VbeMthNy(CurVbe)
End Sub

Private Sub ZZ_VbeMthNyWh()
Brw VbeMthNy(CurVbe)
End Sub

Private Sub ZZ_VbeMthWb()
WbVis VbeMthWb(CurVbe)
End Sub

Private Sub ZZ_VbeMthWs()
WsVis VbeMthWs(CurVbe)
End Sub

Sub ZZ_VbeWsFunNmzDupLines()
'WsVis VbeWsFunNmzDupLines(CurVbe)
End Sub

Function ZZVbeAy() As Vbe()
PushObj ZZVbeAy, CurVbe
Const Fb$ = "C:\Users\user\Desktop\MHD\SAPAccessReports\StockShipRate\StockShipRate\StockShipRate (ver 1.0).accdb"
PushObj ZZVbeAy, AcsOpnFb(Fb).Vbe
End Function

Function CvVbe(A) As Vbe
Set CvVbe = A
End Function
