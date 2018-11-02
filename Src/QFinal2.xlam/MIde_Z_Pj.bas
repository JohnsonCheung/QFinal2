Attribute VB_Name = "MIde_Z_Pj"
Option Explicit
Sub ActPj(PjNm$)
Set CurVbe.ActiveVBProject = Pj(PjNm)
End Sub

Function Pj(A) As VBProject
Set Pj = CurVbe.VBProjects(A)
End Function

Sub PjAct(A As VBProject)
Set CurVbe.ActiveVBProject = A
End Sub

Function PjAddCls(A As VBProject, Nm$) As CodeModule
Set PjAddCls = MdAddOptExpLin(PjAddCmp(A, Nm, vbext_ct_ClassModule).CodeModule)
End Function

Function IsPjNm(A) As Boolean
IsPjNm = AyHas(PjNy, A)
End Function



Sub PjAddClsFmPj(A As VBProject, FmPj As VBProject, ClsNy0)
Dim I, ClsNy$(), ClsAy() As CodeModule
ClsNy = CvNy(ClsNy0)
For Each I In A
    MdCpy CvMd(I), A
Next
End Sub

Function PjAddCmp(A As VBProject, Nm, Ty As vbext_ComponentType) As VBComponent
If PjHasCmp(A, Nm) Then
    Er "PjAddCmp", "[Pj] already has [Cmp]", A.Name, Nm
End If
Set PjAddCmp = A.VBComponents.Add(Ty)
PjAddCmp.Name = Nm
End Function

Function PjAddCmpLines(A As VBProject, Nm, Ty As vbext_ComponentType, Lines$)
Dim O As VBComponent
Set O = PjAddCmp(A, Nm, Ty): If IsNothing(O) Then Stop
MdLinesApp O.CodeModule, Lines
Set PjAddCmpLines = O
End Function

Sub PjAddMdPfx(A As VBProject, B As WhMd, MdPfx$)
Dim Md As CodeModule, M
For Each M In AyNz(PjMdAy(A, B))
    Set Md = M
    MdRen Md, MdPfx & MdNm(Md)
Next
End Sub

Function PjAddMod(A As VBProject, Nm) As CodeModule
Set PjAddMod = MdAddOptExpLin(PjAddCmp(A, Nm, vbext_ct_StdModule).CodeModule)
End Function

Sub PjAddRfFfnAy(A As VBProject, RfFfnAy$())
Dim F
For Each F In RfFfnAy
    If Not PjHasRfFfn(A, CStr(F)) Then
        A.References.AddFromFile F
    End If
Next
End Sub

Function PjClsAndMdNy(A As VBProject, Optional Patn$, Optional Exl$) As String()
'PjClsAndMdNy = PjCmpNy(A, WhMd("Cls Mod", WhNm(Patn, Exl)))
End Function

Function PjClsAy(A As VBProject, Optional B As WhNm) As CodeModule()
PjClsAy = PjMdAy(A, WhMd("Cls"))
End Function

Function PjClsNy(A As VBProject, Optional B As WhNm) As String()
PjClsNy = PjCmpNy(A, WhMd("Cls", B))
End Function

Sub Z_PjClsNy()
AyDmp PjClsNy(CurPj)
End Sub

Function PjCmp(A As VBProject, Nm) As VBComponent
Set PjCmp = A.VBComponents(Nm)
End Function

Function PjCmpAy(A As VBProject, Optional B As WhMd) As VBComponent()
Dim Cmp
For Each Cmp In AyNz(PjMdAy(A, B))
    PushObj PjCmpAy, Cmp
Next
End Function

Function PjCmpNy(A As VBProject, Optional B As WhMd) As String()
PjCmpNy = ItrNy(PjCmpAy(A, B))
End Function

Sub PjCompile(A As VBProject)
PjGo A
AssCompileBtn PjNm(A)
With CompileBtn
    If .Enabled Then
        .Execute
        Debug.Print PjNm(A), "<--- Compiled"
    Else
        Debug.Print PjNm(A), "already Compiled"
    End If
End With
TileVBtn.Execute
SavBtn.Execute
End Sub

Sub PjCpyToSrc(A As VBProject)
FfnCpyToPth A.Filename, PjSrcPth(A), OvrWrt:=True
End Sub

Sub PjCpyToSrcPth(A As VBProject)
FfnCpyToPth A.Filename, PjSrcPth(A), OvrWrt:=True
End Sub

Sub PjCrtCmp(A As VBProject, Nm, Ty As vbext_ComponentType)
Dim O As VBComponent
Set O = A.VBComponents.Add(Ty)
O.Name = Nm
End Sub

Sub PjCrtMd(A As VBProject, MdNm$)
PjCrtCmp A, MdNm, vbext_ct_StdModule
End Sub

Private Sub Z_PjCurPjx()
Ass CurPj.Name = "lib1"
End Sub

Sub PjDltMd(A As VBProject, MdNm$)
If Not PjHasMd(A, MdNm) Then Exit Sub
A.VBComponents.Remove A.VBComponents(MdNm)
End Sub

Function PjEnsCls(A As VBProject, ClsNm$) As CodeModule
Set PjEnsCls = PjEnsCmp(A, ClsNm, vbext_ct_ClassModule)
End Function

Function PjEnsCmp(A As VBProject, Nm, Optional Ty As vbext_ComponentType = vbext_ct_StdModule) As CodeModule
If Not PjHasCmp(A, Nm) Then
    PjCrtCmp A, Nm, Ty
End If
Set PjEnsCmp = A.VBComponents(Nm).CodeModule
End Function

Function PjEnsMod(A As VBProject, MdNm) As CodeModule
Set PjEnsMod = PjEnsCmp(A, MdNm, vbext_ct_StdModule)
End Function

Function PjEnsStd(A As VBProject, StdNm$) As CodeModule
Set PjEnsStd = PjEnsCmp(A, StdNm, vbext_ct_StdModule)
End Function

Sub PjEnsZDashMthAsPrv(A As VBProject)
ItrDo PjMdAy(A), "MdEnsZ3DMthAsPrivate"
End Sub

Sub PjEnsZZDashAsPrv(A As VBProject)

End Sub

Sub PjEnsZZDashAsPub(A As VBProject)
AyDo PjMdAy(A), "MdEnsZZDashAsPrv"
End Sub

Sub PjExp(A As VBProject)
PjExpSrc A
PjExpRf A
End Sub

Sub PjExpRf(A As VBProject)
Ass Not PjIsUnderSrcPth(A)
AyWrt PjRfLy(A), PjRfCfgFfn(A)
End Sub

Sub PjExpSrc(A As VBProject)
PjCpyToSrc A
PthClrFil PjSrcPth(A)
Dim Md As CodeModule, I
For Each I In PjModAy(A)
    Set Md = I
    MdExp Md
Next
End Sub

Sub PjExport(A As VBProject)
Debug.Print "PjExport: " & PjNm(A) & "-----------------------------"
Dim P$: P = PjSrcPth(A)
If P = "" Then
    Debug.Print FmtQQ("PjExport: Pj(?) does not have FileName", A.Name)
    Exit Sub
End If
PthClrFil P 'Clr SrcPth ---
FfnCpyToPth A.Filename, P, OvrWrt:=True
Dim I, Ay() As CodeModule
Ay = PjMdAy(A)
If Sz(Ay) = 0 Then Exit Sub
For Each I In Ay
    MdExport CvMd(I)  'Exp each md --
Next
'AyWrt PjRfLy(A), PjRfCfgFfn(A) 'Exp rf -----
End Sub

Function PjFfn$(A As VBProject)
On Error Resume Next
PjFfn = A.Filename
End Function

Function PjFfnApp(PjFfn) ' Return either Xls.Application (CurXls) or Acs.Application (Function-static)
Static Y As New Access.Application
Select Case True
Case IsFxa(PjFfn): FxaOpn PjFfn: Set PjFfnApp = CurXls
Case IsFb(PjFfn): Y.OpenCurrentDatabase PjFfn: Set PjFfnApp = Y
Case Else: Stop
End Select
End Function


Function PjFn$(A As VBProject)
PjFn = FfnFn(PjFfn(A))
End Function

Function PjFstMbr(A As VBProject) As CodeModule
Dim Cmp As VBComponent
For Each Cmp In A.VBComponents
    Set PjFstMbr = Cmp.CodeModule
    Exit Function
Next
End Function

Function CvPj(I) As VBProject
Set CvPj = I
End Function

Function PjFstMd(A As VBProject) As CodeModule
Dim Cmp As VBComponent
For Each Cmp In A.VBComponents
    If Cmp.Type = vbext_ct_StdModule Then
        Set PjFstMd = Cmp.CodeModule
        Exit Function
    End If
Next
End Function

Function PjFunBdyDic(A As VBProject) As Dictionary
Stop '
End Function

Function PjFunPfxAy(A As VBProject) As String()
Dim Ay() As CodeModule: Ay = PjMdAy(A)
Dim Ay1(): Ay1 = AyMap(Ay, "MdFunPfx")
PjFunPfxAy = AyFlat(Ay1)
End Function

Sub PjGo(A As VBProject)
ClsAllWin
Dim Md As CodeModule
Set Md = PjFstMbr(A)
If IsNothing(Md) Then
    Exit Sub
End If
Md.CodePane.Show
TileVBtn.Execute
DoEvents
End Sub

Sub PjGoMdNm(A As VBProject, MdNm$)
ClsAllWinExl ApWinAy(MdWin(Md(MdNm)))
End Sub

Function PjHasCmp(A As VBProject, Nm) As Boolean
PjHasCmp = ItrHasNm(A.VBComponents, Nm)
End Function

Function PjHasCmpWhRe(A As VBProject, Re As RegExp) As Boolean
PjHasCmpWhRe = ItrHasNmWhRe(A.VBComponents, Re)
End Function

Function PjHasMd(A As VBProject, Nm) As Boolean
Dim T As vbext_ComponentType
If Not ItrHasNm(A.VBComponents, Nm) Then Exit Function
T = PjCmp(A, Nm).Type
If T = vbext_ct_StdModule Then PjHasMd = True: Exit Function
Debug.Print "PjHasMd: Pj(?) has Mbr(?), but it is not Md, but CmpTy(?)", PjNm(A), Nm, CmpTyStr(T)
End Function

Function PjHasNoStdClsMd(A As VBProject) As Boolean
Dim C As VBComponent
For Each C In A.VBComponents
    If C.Type = vbext_ComponentType.vbext_ct_ClassModule Then Exit Function
    If C.Type = vbext_ComponentType.vbext_ct_StdModule Then Exit Function
Next
PjHasNoStdClsMd = True
End Function

Function PjHasRf(A As VBProject, RfNm)
Dim RF As VBIDE.Reference
For Each RF In A.References
    If RF.Name = RfNm Then PjHasRf = True: Exit Function
Next
End Function

Function PjHasRfFfn(A As VBProject, RfFfn) As Boolean
Dim R As Reference
For Each R In A.References
    If R.FullPath = RfFfn Then PjHasRfFfn = True: Exit Function
Next
End Function

Function PjHasRfNm(A As VBProject, RfNm$) As Boolean
Dim I, R As Reference
For Each I In A.References
    Set R = I
    If R.Name = RfNm Then PjHasRfNm = True: Exit Function
Next
End Function

Sub PjImpRf(A As VBProject, RfCfgPth$)
Dim B As Dictionary: Set B = FtDic(RfCfgPth & "PjRf.Cfg")
Dim K
For Each K In B.Keys
    PjAddRf A, K, B(K)
Next
End Sub

Sub PjImpSrcFfn(A As VBProject, SrcFfn)
A.VBComponents.Import SrcFfn
End Sub

Function PjIsUnderSrcPth(A As VBProject) As Boolean
Dim B$: B = PjPth(A)
If PthFdr(B) = "Src" Then Stop
End Function

Function PjIsUsrLib(A As VBProject) As Boolean
PjIsUsrLib = PjIsFxa(A)
End Function

Function PjMd(A As VBProject, Nm) As CodeModule
Set PjMd = PjCmp(A, Nm).CodeModule
End Function

Function PjMdAy(A As VBProject, Optional B As WhMd) As CodeModule()
If IsNothing(B) Then
    PjMdAy = ItrPrpAyInto(A.VBComponents, "CodeModule", PjMdAy)
    Exit Function
End If
Dim C
For Each C In AyNz(ItrWhNm(A.VBComponents, B.Nm))
    With CvCmp(C)
        If AySel(B.InCmpTy, .Type) Then
            PushObj PjMdAy, .CodeModule
        End If
    End With
Next
End Function

Private Sub Z_PjMdAy()
Dim O() As CodeModule
O = PjMdAy(CurPj)
Dim I, Md As CodeModule
For Each I In O
    Set Md = I
    Debug.Print MdNm(Md)
Next
End Sub

Sub PjMdDicApp(A As VBProject, MdDic As Dictionary)
Dim MdNm
For Each MdNm In MdDic.Keys
    PjEnsMod A, MdNm
    MdLinesApp PjMd(A, MdNm), MdDic(MdNm)
Next
End Sub

Function PjMdNy(A As VBProject, Optional B As WhMd) As String()
PjMdNy = PjCmpNy(A, B)
End Function

Function PjMdNy_With_TstSub(A As VBProject) As String()
Dim I As VBComponent
Dim O$()
For Each I In A.VBComponents
    If I.Type = vbext_ct_StdModule Then
        If MdHasTstSub(I.CodeModule) Then
            Push O, I.Name
        End If
    End If
Next
PjMdNy_With_TstSub = O
End Function

Sub Z_PjMdNy()
AyDmp PjMdNy(CurPj)
End Sub

Function PjMdOpt(A As VBProject, Nm) As CodeModule
If Not PjHasMd(A, Nm) Then Exit Function
Set PjMdOpt = PjMd(A, Nm)
End Function

Function PjModAy(A As VBProject, Optional B As WhNm) As CodeModule()
PjModAy = PjMdAy(A, WhMd("Mod", B))
End Function

Function PjModClsNy(A As VBProject, Optional B As WhNm) As String()
PjModClsNy = PjCmpNy(A, WhMd("Mod Cls", B))
End Function

Function PjModNy(A As VBProject, Optional B As WhNm) As String()
PjModNy = PjCmpNy(A, WhMd("Mod", B))
End Function

Function PjMthKy(A As VBProject, Optional IsWrap As Boolean) As String()
PjMthKy = AyMapPXSy(PjMdAy(A), "MdMthKy", IsWrap)
End Function

Function PjMthKySq(A As VBProject) As Variant()
PjMthKySq = MthKy_Sq(PjMthKy(A, True))
End Function

Function PjMthKyWs(A As VBProject) As Worksheet
Set PjMthKyWs = WsVis(SqWs(PjMthKySq(A)))
End Function

Function PjMthLinDry(A As VBProject) As Variant()
Dim M
For Each M In AyNz(PjMdAy(A))
    PushAy PjMthLinDry, MdMthLinDry(CvMd(M))
Next
End Function

Function PjMthLinDryWP(A As VBProject) As Variant()
Dim M
For Each M In AyNz(PjMdAy(A))
    PushIAy PjMthLinDryWP, MdMthLinDryWP(CvMd(M))
Next
End Function

Private Sub Z_PjMthLinDry()
Dim A(): A = PjMthLinDry(CurPj)
Stop
End Sub

Function PjMthNy(A As VBProject, Optional B As WhMdMth) As String()
Dim Md As CodeModule, I, N$, Ny$()
N = A.Name & "."
For Each I In AyNz(PjMdAy(A, WhMdMthMd(B)))
    Set Md = I
    Ny = MthDDNyWh(MdMthDDNy(Md), B.Mth)
    Ny = AyAddPfx(Ny, N & MdNm(Md) & ".")
    PushAyNoDup PjMthNy, Ny
Next
End Function

Function PjMthSq(A As VBProject) As Variant()
PjMthSq = MthKy_Sq(PjMthKy(A, True))
End Function

Function PjMthWs(A As CodeModule) As Worksheet
Set PjMthWs = WsVis(SqWs(PjMthSq(A)))
End Function

Function PjNm$(A As VBProject)
PjNm = A.Name
End Function

Function PjNy() As String()
PjNy = ItrNy(CurVbe.VBProjects)
End Function

Function PjPatnLy(A As VBProject, Patn$) As String()
Dim I, Md As CodeModule, O$()
For Each I In PjMdAy(A)
   Set Md = I
   PushAy O, MdPatnLy(Md, Patn)
Next
PjPatnLy = O
End Function

Function PjPjPrpInfDt(A As VBProject) As Dt

End Function

Function PjPth$(A As VBProject)
PjPth = FfnPth(A.Filename)
End Function

Function PjReadRfCfg(A As VBProject) As String()
Const CSub$ = "PjReadRfCfg"
Dim B$: B = PjRfCfgFfn(A)
If Not FfnIsExist(B) Then Er CSub, "{Pj-Rf-Cfg-Fil} not found", B
PjReadRfCfg = FtLy(B)
End Function

Sub PjRenMdByPfx(A As VBProject, FmMdPfx$, ToMdPfx$)
Dim CvNy$()
Dim Ny$()
'    Ny = PjMdNy(A, "^" & FmMdPfx)
    CvNy = AyMapAsgSy(Ny, "RplPfx", FmMdPfx, ToMdPfx)
Dim MdAy() As CodeModule
    Dim MdNm
    Dim Md As CodeModule
    For Each MdNm In Ny
        Set Md = PjMd(A, CStr(MdNm))
        PushObj MdAy, Md
    Next
Dim I%, U%
    For I = 0 To UB(CvNy)
        MdRen MdAy(I), CvNy(I)
    Next
End Sub

Private Sub Z_PjRenMdByPfx()
PjRenMdByPfx CurPj, "A_", ""
End Sub
Sub PjRmvMdNmPfx(A As VBProject, Pfx$)
Dim I
For Each I In PjMdAy(A, WhMd(Nm:=WhNm("^" & Pfx)))
    MdRmvNmPfx CvMd(I), Pfx
Next
End Sub

Sub PjRmvMdPfx(A As VBProject, B As WhMd, MdPfx$)
Dim Md As CodeModule, M
For Each M In AyNz(PjMdAy(A, B))
    Set Md = M
    Md.Parent.Name = RmvPfx(MdNm(A), MdPfx)
Next
End Sub

Sub PjRmvOptCmpDbLin(A As VBProject)
Dim I
For Each I In PjMdAy(A)
   MdRmvOptCmpDb CvMd(I)
Next
End Sub

Sub PjRmvRf(A As VBProject, RfNy0$)
AyDoPX CvNy(RfNy0), "PjRmvRf__X", A
PjSav A
End Sub

Private Sub PjRmvRf__X(A As VBProject, RfNm$)
If PjHasRfNm(A, RfNm) Then
    Debug.Print FmtQQ("PjAddRf: Pj(?) already has RfNm(?)", A.Name, RfNm)
    Exit Sub
End If
Dim RfFfn$: RfFfn = PjRfNmRfFfn(A, RfNm)
If PjHasRfFfn(A, RfFfn) Then
    Debug.Print FmtQQ("PjAddRf: Pj(?) already has RfFfnNm(?)", A.Name, RfFfn)
    Exit Sub
End If
A.References.AddFromFile RfFfn
End Sub

Sub PjSav(A As VBProject)
If FstChr(PjNm(A)) <> "Q" Then
    Debug.Print "PjSav: Project Name begin with Q, it is not saved: "; PjNm(A)
    Exit Sub
End If
If A.Saved Then
    Debug.Print FmtQQ("PjSav: Pj(?) is already saved", A.Name)
    Exit Sub
End If
Dim Fn$: Fn = PjFn(A)
If Fn = "" Then
    Debug.Print FmtQQ("PjSav: Pj(?) needs saved first", A.Name)
    Exit Sub
End If
PjAct A
If ObjPtr(CurPj) <> ObjPtr(A) Then Stop: Exit Sub
Dim B As CommandBarButton: Set B = SavBtn
If Not StrIsEq(B.Caption, "&Save " & Fn) Then Stop
B.Execute
If A.Saved Then Stop
Debug.Print FmtQQ("PjSav: Pj(?) is saved <---------------", A.Name)
End Sub

Function PjSrc(A As VBProject) As String()
Dim C As VBComponent
For Each C In A.VBComponents
    PushAy PjSrc, MdSrc(C.CodeModule)
Next
End Function

Function PjSrcPth$(A As VBProject)
Dim Ffn$: Ffn = PjFfn(A)
Dim Fn$: Fn = FfnFn(Ffn)
Dim O$:
O = FfnPth(A.Filename) & "Src\": PthEns O
O = O & Fn & "\":                PthEns O
PjSrcPth = O
End Function

Sub PjSrcPthBrw(A As VBProject)
PthBrw PjSrcPth(A)
End Sub

Function PjTim(A As VBProject) As Date
PjTim = FfnTim(PjFfn(A))
End Function

Function Pj_ClsNy_With_TstSub(A As VBProject) As String()
Dim I As VBComponent
Dim O$()
For Each I In A.VBComponents
    If I.Type = vbext_ct_ClassModule Then
        If MdHasTstSub(I.CodeModule) Then
            Push O, I.Name
        End If
    End If
Next
Pj_ClsNy_With_TstSub = O
End Function

Sub Pj_Gen_TstClass(A As VBProject)
If PjHasCmp(A, "Tst") Then
    CmpRmv PjCmp(A, "Tst")
End If
PjAddCls A, "Tst"
PjMd(A, "Tst").AddFromString Pj_TstClass_Bdy(A)
End Sub

Function Pj_TstClass_Bdy$(A As VBProject)
Dim N1$() ' All Class Ny with 'Friend Sub Z' method
Dim N2$()
Dim A1$, A2$
Const Q1$ = "Sub ?()|Dim A As New ?: A.Z|End Sub"
Const Q2$ = "Sub ?()|#.?.Z|End Sub"
N1 = Pj_ClsNy_With_TstSub(A)
A1 = SeedExpand(Q1, N1)
N2 = PjMdNy_With_TstSub(A)
A2 = Replace(SeedExpand(Q2, N2), "#", A.Name)
Pj_TstClass_Bdy = A1 & vbCrLf & A2
End Function

Sub Z()
Z_PjMdDicApp
End Sub

Sub ZZ_PjCompile()
PjCompile CurPj
End Sub

Private Sub ZZ_PjHasMd()
Ass PjHasMd(CurPj, "Drs") = False
Ass PjHasMd(CurPj, "A__Tool") = True
End Sub

Private Sub ZZ_PjSav()
PjSav CurPj
End Sub

Private Sub ZZ_PjSrtCmpRptWb()
Dim O As Workbook: Set O = PjSrtCmpRptWb(CurPj, Vis:=True)
Stop
End Sub

Private Sub Z_PjMdDicApp()
Dim MdDic As New Dictionary
Dim ToPj As VBProject: Set ToPj = TmpPj
PjMdDicApp ToPj, MdDic
End Sub
