Attribute VB_Name = "MDao_Lnk_Import"
Option Explicit
Sub Z()
Z_ApScl
Z_DbAddTd
Z_DbtfAddExpr
Z_DrsEnsRplDbt
Z_DrsInsUpdDbt
Z_LblSeqAy
Z_PfxSsl_Sy
Z_SclChk
End Sub
Sub AssAyIsSamSz(A, B)
Ass Sz(A) = Sz(B)
End Sub

Sub AssDicHasKeyss(A As Dictionary, Keyss$)
AssDicHasKy A, SslSy(Keyss)
End Sub

Sub AssEqDic(A As Dictionary, B As Dictionary)
If Not IsEqDic(A, B) Then Stop
End Sub

Private Sub AssFil()
'FilLinAyAss ActFilLinAy
End Sub

Function AssNegEle(A)
Dim I, J&, O$()
For Each I In AyNz(A)
    If I < 0 Then
        PushI O, J & ": " & I
        J = J + 1
    End If
Next
If Sz(O) > 0 Then
    Er "AssNegEle", "In [Ay], there are [negative-element (Ix Ele)]", A, O
End If
End Function

Sub AssNoFfn(Ffn, FilKd$, CSub$)
If FfnIsExist(Ffn) Then Er "AssNoFfn", "Given [" & FilKd & "] exist", Ffn
End Sub

Function AyFstRmvTT$(A, T1$, T2$)
Dim X, X1$, X2$, Rst$
For Each X In AyNz(A)
    LinAsg2TRst X, X1, X2, Rst
    If X1 = T1 Then
        If X2 = T2 Then
            AyFstRmvTT = X
            Exit Function
        End If
    End If
Next
End Function

Private Function C_WhyBlkIsEr_MsgAy() As String()
Dim O$()
Push O, "The block is error because, it is none of these [RmkBlk SqBlk PrmBlk SwBlk]"
Push O, "SwBlk is all remark line or SwLin, which is started with ?"
Push O, "PrmBlk is all remark line or PrmLin, which is started with %"
Push O, "SqBlk is first non-remark begins with these [sel seldis drp upd] with optionally ?-Pffx"
Push O, "RmkBlk is all remark lines"
End Function


Function CurFb$()
On Error Resume Next
CurFb = CurDb.Name
End Function

Function CvIxl(A) As Lnx
Set CvIxl = A
End Function

Function DicKyJnVal$(A As Dictionary, Ky, Optional Sep$ = vbCrLf & vbCrLf)
Dim O$(), K
For Each K In AyNz(Ky)
    If A.Exists(K) Then
        PushI O, A(K)
    End If
Next
DicKyJnVal = Join(O, Sep)
End Function

Function DSpecNm$(A)
DSpecNm = TakAftDotOrAll(LinT1(A))
End Function

Function FfnNotFoundChk(FfnAy0) As String()
Dim I, O$()
For Each I In CvFfnAy(FfnAy0)
    If Not FfnIsExist(CStr(I)) Then
        PushI O, I
    End If
Next
If Sz(O) = 0 Then Exit Function
FfnNotFoundChk = MsgAp_Ly("[File(s)] not found", O)
End Function

Function FnySubIxAy(A$(), SubFny0) As Integer()
Dim SubFny$(): SubFny = CvNy(SubFny0)
If Sz(SubFny) = 0 Then Stop
Dim O%(), U&, J%
U = UB(SubFny)
ReSz O, U
For J = 0 To U
    O(J) = AyIx(A, SubFny(J))
    If O(J) = -1 Then Stop
Next
FnySubIxAy = O
End Function

Sub FnyIxAsg(Fny$(), FldLvs$, ParamArray OAp())
'FnyIxAsg=Field Index Array
Dim A$(): A = SplitSpc(FldLvs)
Dim I&(): I = AyIxAy(Fny, A)
Dim J%
For J = 0 To UB(I)
    OAp(J) = I(J)
Next
End Sub

Function FnyAlignQuote(Fny$()) As String()
FnyAlignQuote = AyAlignL(AySqBktQuoteIfNeed(Fny))
End Function

Sub FnyWhFldLvs(Fny$(), FldLvs$, ParamArray OAp())
'FnyWhFldLvs=Field Index Array
Dim A$(): A = SplitSpc(FldLvs)
Dim I&(): I = AyIxAy(Fny, A)
Dim J%
For J = 0 To UB(I)
    OAp(J) = I(J)
Next
End Sub


Sub FSharpBuildKnownColor()
'// Learn more about F# at http://fsharp.org
'// See the 'F# Tutorial' project for more help.
'open System.Drawing
'open System
'open System.IO
'open System.Windows.Forms
'
'type slis = String list
'type sy = String[]
'type sseq = String seq
'let slis_lines(a:slis) = String.Join("\r\n",a)
'let sy_lines(a:sy) = String.Join("\r\n",a)
'let str_wrt ft a = File.WriteAllText(ft,a)
'let sseq_wrt ft (a:sseq) = File.WriteAllLines(ft,a)
'let slis_wrt ft a = a|> sseq_wrt ft
'let mayStr_wrt a ft = match a with | Some a -> str_wrt a ft | _ -> ()
'Let colorConstFt = "ColorLines.Txt"
'//let knownColor_lin a = a.ToString() + " " + Color.FromKnownColor(a).ToArgb().ToString()
'let knownColor_lin a = "Const " + a.ToString() + "& = " + Color.FromKnownColor(a).ToArgb().ToString()
'let sy_wrt a ft = a |> sseq_wrt ft
'let arr_seq<'a>(a:Array) = seq { for i in a -> unbox i }
'let arr_ay<'a>(a:Array) = [|for i in a -> unbox i|]
'let arr_lis<'a>(a:Array) = [for i in a -> unbox i]
'let knownColorArr = Enum.GetValues(KnownColor.ActiveBorder.GetType())
'let knownColorLis = knownColorArr |> arr_lis<KnownColor>
'let colorConstLis = knownColorLis |> List.map knownColor_lin |> List.sort
'let wrt_colorConstFt() = slis_wrt colorConstFt colorConstLis
'[<EntryPoint>]
'let main argv =
'    printfn "%A" argv
'//    MessageBox.Show System.Environment.CurrentDirectory |> ignore
'    do wrt_colorConstFt()
'    0 // return an integer exit code
End Sub
Private Sub HowToEnsFirstTime_FmtSpec()
'No table-Spec
'No rec-Fmt
End Sub

Private Function InActInpy(ASw() As LnkASw, FmSw() As LnkFmSw) As String()
Dim O$(), I, IFm As LnkFmSw, SwNm$, TF As Boolean
'If Sz(ASw) = 0 Then Exit Function

'For Each I In FmSw
'    Set IFm = I
'    SwNm = IFm.SwNm
'    TF = IFm.TF
'    If Not InActInpy__zSel(SwNm, TF, ASw) Then
'        PushAy O, SslSy(Rmv3T(IFm.Inpy))
'    End If
'Next
InActInpy = O
End Function

Private Function InActInpy__zSel(SwNm$, TF As Boolean, ASw() As LnkASw) As Boolean
'Dim IA As LnkASw, I
'For Each I In ASw
'    Set IA = I
'    If SwNm = IA.SwNm Then
'        InActInpy__zSel = IA.TF = TF
'        Exit Function
'    End If
'Next
'Stop
End Function

Function ISpecINm$(A$)
ISpecINm = LinT1(A)
End Function


Sub IxAyAss(IxAy, U)
Dim O(), Ix
For Each Ix In AyNz(IxAy)
    If 0 > Ix Or Ix > U Then PushI O, Ix
Next
If Sz(O) > 0 Then
    Er "IxAyAss", "[Er-Ix] found in [IxAy], where [U]", O, IxAy, U
End If
End Sub

Function IxAyIsAllGE0(IxAy) As Boolean
Dim Ix
For Each Ix In AyNz(IxAy)
    If Ix < 0 Then Exit Function
Next
IxAyIsAllGE0 = True
End Function

Function KyK0_BExpr$(Ky$(), K0)
Dim U%, S$, Vy
U = UB(Ky)
Vy = CvNy(K0)
If U <> UB(Vy) Then Stop
Dim O$(), J%, V
For J = 0 To U
    If IsNull(Vy(J)) Then
        Push O, Ky(J) & " is null"
    Else
        V = Vy(J): GoSub GetS
        Push O, Ky(J) & "=" & S
    End If
Next
KyK0_BExpr = Join(O, " and ")
Exit Function
GetS:
Select Case True
Case IsStr(V): S = "'" & V & "'"
Case IsDate(V): S = "#" & V & "#"
Case IsBool(V): S = IIf(V, "TRUE", "FALSE")
Case IsNumeric(V): S = V
Case Else: Stop
End Select
Return
End Function

Function MapStrDic(A) As Dictionary
Set MapStrDic = S1S2AyStrDic(A)
End Function

Function MapVbl_Dic(A) As Dictionary
Dim Ay$(): Ay = SplitVBar(A)
Dim O As New Dictionary
Dim I
If Sz(Ay) > 0 Then
    For Each I In Ay
        With BrkBoth(I, ":")
            O.Add .S1, .S2
        End With
    Next
End If
Set MapVbl_Dic = O
End Function

Function MayDicVal(MayDic As Dictionary, K)
If Not IsNothing(MayDic) Then MayDicVal = DicVal(MayDic, K)
End Function

Function MissingFnyChk(MissFny$(), ExistingFny$(), A As Database, T) As String()
If Sz(MissFny) = 0 Then Exit Function
Dim LnkVbl$
LnkVbl = DbtLnkVbl(A, T)
Const F1$ = "Excel file       : "
Const T1$ = "Worksheet        : "
Const C1$ = "Worksheet column : "
Const F2$ = "Database file: "
Const T2$ = "Table        : "
Const C2$ = "Field        : "
Dim X$, Y$, Z$, F0$, C0$, T0$
    AyAsg SplitVBar(LnkVbl), X, Y, Z
    Select Case X
    Case "LnkFb", "Lcl"
        F0 = F1
        T0 = T1
        C0 = C1
    Case "LnkFx"
        F0 = F2
        T0 = T2
        C0 = C2
    Case Else: Stop
    End Select
Dim O$()
    Dim I
    Push O, F0 & Y
    Push O, T0 & Z
    PushUnderLin O
    For Each I In ExistingFny
        Push O, C0 & QuoteSqBkt(CStr(I))
    Next
    PushUnderLin O
    For Each I In MissFny
        Push O, C0 & QuoteSqBkt(CStr(I))
    Next
'    PushMsgUnderLinDbl O, FmtQQ("Above ? are missing", D)
Stop '
MissingFnyChk = O
End Function

Sub NDriveMap()
NDriveRmv
Shell "Subst N: c:\users\user\desktop\MHD"
End Sub

Sub NDriveRmv()
Shell "Subst /d N:"
End Sub

Sub Never()
Const CSub$ = "Never"
Er CSub, "Should never reach here"
End Sub


Function NewIntSeq(N&, Optional IsFmOne As Boolean) As Integer()
Dim O%(): ReDim O(N - 1)
Dim J&
If IsFmOne Then
    For J = 0 To N - 1
        O(J) = J + 1
    Next
Else
    For J = 0 To N - 1
        O(J) = J
    Next
End If
NewIntSeq = O
End Function

Function NewIpFil(Ly$()) As LnkIpFil()
If Sz(Ly) = 0 Then Exit Function
Dim O() As LnkIpFil, J%, L, Ay
ReDim O(UB(Ly))
For Each L In Ly
'    Ay = AyT1Rst(SslSy(L))
    Stop '
'    Set O(J) = NewLnkIpFil(L)
    O(J).Fil = Ay(0)
    'O(J).Inpy = CvSy(Ay(1))
    J = J + 1
Next
NewIpFil = O
End Function

Function NewSimTy(SimTyStr$) As eSimTy
Dim O As eSimTy
Select Case UCase(SimTyStr)
Case "TXT": O = eTxt
Case "NBR": O = eNbr
Case "LGC": O = eLgc
Case "DTE": O = eDte
Case Else: O = eOth
End Select
NewSimTy = O
End Function

Function NewStExt(Lin) As LnkStExt
'Dim O As New LnkStExt
'With O
'    AyAsg Lin3TAy(Lin), .LikInp, .F, , .Ext
'End With
'Set NewStExt = O
End Function

Private Function NewStFld(Lin) As LnkStFld
'Dim O As New LnkStFld, A$
'With O
'    AyAsg Lin2TAy(Lin), .Stu, , A
'    .Fny = SslSy(A)
'End With
'Set NewStFld = O
End Function

Function NewSy(U&) As String()
Dim O$()
If U > 0 Then ReDim O(U)
NewSy = O
End Function

Function NewTd(T, FdAy() As DAO.Field2) As DAO.TableDef
Dim O As New DAO.TableDef, F
O.Name = T
For Each F In FdAy
    O.Fields.Append F
Next
Set NewTd = O
End Function

Function NmNxtSeqNm$(A, Optional NDig% = 3) _
'Nm-A can be XXX or XXX_nn
'   If XXX, return XXX_001   '<-- # of zero depends on NDig
'   If XXX_nn, return XXX_mm '<-- mm is nn+1, # of digit of nn and mm depends on NDig
If NDig = 0 Then Stop
Dim R$
    R = Right(A, NDig + 1)

If Left(R, 1) <> "_" Then GoTo Case1
If Not IsNumeric(Mid(R, 2)) Then GoTo Case1

Dim L$: L = Left(A, Len(A) - NDig)
Dim Nxt%: Nxt = Val(Mid(R, 2)) + 1
NmNxtSeqNm = Left(A, Len(A) - NDig) + Pad0(Nxt, NDig)
Exit Function

Case1:
    NmNxtSeqNm = A & "_" & StrDup(NDig - 1, "0") & "1"
End Function

Function Ny0SqBktCsv$(A)
Dim B$(), C$()
B = CvNy(A)
C = AyQuoteSqBkt(B)
Ny0SqBktCsv = JnComma(C)
End Function

Private Function NyEy(Ny$(), A() As LnkStEle) As String()

End Function

Function NyEySqy$(Ny$(), Ey$())
AyAssSamSz Ny, Ey
If IsEqAy(Ny, Ey) Then
    NyEySqy = "Select" & vbCrLf & "    " & JnComma(Ny)
    Exit Function
End If
Dim N$()
    N = AyAlignL(Ny)
Dim E$()
    Dim J%
    E = Ey
    For J = 0 To UB(E)
        If E(J) <> "" Then E(J) = QuoteSqBkt(E(J))
    Next
    E = AyAlignL(E)
    For J = 0 To UB(E)
        If Trim(E(J)) <> "" Then E(J) = E(J) & " As "
    Next
    E = AyAlignL(E)
Dim O$()
    O = AyabAdd(E, N)
NyEySqy = Join(O, "," & vbCrLf)
End Function

Function NyLnxAy(Ny0) As String()
'It is to return 2 lines with
'first line is 0   1     2 ..., where 0,1,2.. are ix of A$()
'second line is each element of A$() separated by A
'Eg, A$() = "A BBBB CCC DD"
'return 2 lines of
'0 1    2   3
'A BBBB CCC DD
Dim Ny$(): Ny = CvNy(Ny0)
If Sz(Ny) = 0 Then Exit Function
Dim A1$()
Dim A2$()
Dim U&: U = UB(Ny)
ReSz A1, U
ReSz A2, U
Dim O$(), J%, L$, W%
For J = 0 To U
    L = Len(Ny(J))
    W = Max(L, Len(J))
    A1(J) = AlignL(J, W)
    A2(J) = AlignL(Ny(J), W)
Next
Push O, JnSpc(A1)
Push O, JnSpc(A2)
NyLnxAy = O
End Function

Function OkShow(Ok$()) As String()
OkShow = SyShow("Ok", Ok)
End Function

Function OupPth$()
OupPth = PthEns(CurDbPth & "Output\")
End Function

Function OupPthPm$()
OupPthPm = PnmVal("OupPth")
End Function

Function PfxSsl_Sy(A) As String()
Dim Ay$(), Pfx$
Ay = SslSy(A)
Pfx = AyShf(Ay)
PfxSsl_Sy = AyAddPfx(Ay, Pfx)
End Function

Function PgmDb_DtaDb(A As Database) As Database
Set PgmDb_DtaDb = DBEngine.OpenDatabase(PgmDb_DtaFb(A))
End Function

Function PgmDb_DtaFb$(A As Database)

End Function

Function PgmObjPth$()
PgmObjPth = PthEns(CurDbPth & "PgmObj\")
End Function

Function PgmPth$()
PgmPth = FfnPth(Excel.Application.Vbe.ActiveVBProject.Filename)
End Function

Function PnmFfn$(A)
PnmFfn = PnmPth(A) & PnmFn(A)
End Function

Function PnmFn$(A)
PnmFn = PnmVal(A & "Fn")
End Function

Function PnmPth$(A)
PnmPth = PthEnsSfx(PnmVal(A & "Pth"))
End Function

Property Get PnmVal$(Pnm$)
PnmVal = CurDb.TableDefs("Prm").OpenRecordset.Fields(Pnm).Value
End Property

Property Let PnmVal(Pnm$, V$)
Stop
'Should not use
With CurDb.TableDefs("Prm").OpenRecordset
    .Edit
    .Fields(Pnm).Value = V
    .Update
End With
End Property

Function Prepend$(S, P)
If S <> "" Then Prepend = P & S
End Function

Function PrependDash$(S)
PrependDash = Prepend(S, "-")
End Function

Sub PrmBrw()
RsBrw TblRs("Prm")
End Sub

Function PrmDotLin$(A$)
Dim ArgAy$()
ArgAy = SplitComma(A)
ArgAy = AyMapSy(ArgAy, "ArgStr")
End Function

Function ProdPth$()
ProdPth = "N:\SAPAccessReports\"
End Function

Sub ResClr(A$)
DbResClr CurDb, A
End Sub

Function ReSeqSpec_OLinFldAy(A) As String()
Dim B$()
B = SplitVBar(A)
AyShf B
ReSeqSpec_OLinFldAy = AyTakT1(B)
End Function

Function ReSeqSpec_OutLin(A, F) As Byte
Dim Ay$(), Ssl, J%
Ay = SplitVBar(A)
If SslHas(Ay(0), F) Then Exit Function
For J = 1 To UB(Ay)
    Select Case SslIx(Ssl, F)
    Case 0: Stop
    Case Is > 0
        ReSeqSpec_OutLin = 2
    End Select
Next
End Function

Function ReSeqSpecFny(A) As String()
Dim Ay$(), D As Dictionary, O$(), L1$, L
Ay = SplitVBar(A)
If Sz(Ay) = 0 Then Exit Function
L1 = AyShf(Ay)
Set D = LyTRst_Dic(Ay)
For Each L In SslSy(L1)
    If D.Exists(L) Then
        Push O, D(L)
    Else
        Push O, L
    End If
Next
ReSeqSpecFny = SslSy(JnSpc(O))
End Function

Function RTrimWhite$(S)
Dim J%
    Dim A$
    For J = Len(S) To 1 Step -1
        If Not IsWhiteChr(Mid(S, J, 1)) Then Exit For
    Next
    If J = 0 Then Exit Function
RTrimWhite = Mid(S, J)
End Function

Private Function SelIntoAy(ActInpy$(), A As LnkSpec) As String()
Dim Inp$, I, J%, O()
ReDim O(UB(ActInpy))
For Each I In ActInpy
'    Set O(J) = New SqlSelInto
    With O(J)
        Inp = I
'        .Ny = InpNy(Inp, A.StInp, A.StFld)
'        .Ey = NyEy(.Ny, A.StEle)
'        .Fm = ">" & Inp
'        .Into = "#I" & Inp
'        .Wh = InpWhBExpr(Inp, A.FmWh)
    End With
    J = J + 1
Next
'SelIntoAy = O
End Function

Sub SelRg_SetXorEmpty(A As Range)
Dim I
For Each I In A
    
Next
End Sub

Function SeqOf__(FmNum, ToNum, OAy)
Dim O&()
ReDim OAy(Abs(FmNum - ToNum))
Dim J&, I&
If ToNum > FmNum Then
    For J = FmNum To ToNum
        OAy(I) = J
        I = I + 1
    Next
Else
    For J = ToNum To FmNum Step -1
        OAy(I) = J
        I = I + 1
    Next
End If
End Function

Function SeqOfInt(FmNum%, ToNum%) As Integer()
SeqOfInt = SeqOf__(FmNum, ToNum, EmpIntAy)
End Function

Function SeqOfLng(FmNum&, ToNum&) As Long()
SeqOfLng = SeqOf__(FmNum, ToNum, EmpLngAy)
End Function

Private Sub SetColr_ToDo()
'TstStep
'   Call Gen
'   Call FmtSpec_Brw 'Edt
'       Edit and Save, then Call Gen will auto import
'where to add autoImp?
'   Under WbFmtAllLo
'AutoImp will show msg if import/noImport
'ColrLy
'   what is the common color name in DotNet Library
'       Use Enums: System.Drawing.KnownColor is no good, because the EnmNm is in seq, it is not return
'       Use VBA.ColorConstants-module is good, but there is few constant
'       Answer: Use *KnownColor to feed in struct-*Color, there is *Color.ToArgb & *KnownColor has name
'               Run the FSharp program.
'               Put the generated file
'                   in
'                       C:\Users\user\Source\Repos\EnumLines\EnumLines\bin\Debug\ColorLines.Const.Txt
'                   Into
'                       C:\Users\user\Desktop\MHD\SAPAccessReports\StockShipRate\StockShipRate\Spec
'               Run ConstGen: It will addd the Const ColorLines = ".... at end
'               Put Fct-Module
'To find some common values to feed into ColrLines
'
'Colr* 4-functions
'    ColrStr_MayColr
'    ColrStr
'    ColrLy
'    ColrLines
End Sub

Sub SetPush(A As Dictionary, K)
If A.Exists(K) Then Exit Sub
A.Add K, Empty
End Sub

Function SetSqpFmt$(Fny$(), Vy())
Dim A$: GoSub X_A
SetSqpFmt = vbCrLf & "  Set" & vbCrLf & A
Exit Function
X_A:
    Dim L$(): L = FnyAlignQuote(Fny)
    Dim R$(): GoSub X_R
    Dim J%, O$(), S$
    S = Space(4)
    For J = 0 To UB(L)
        Push O, S & L(J) & "= " & R(J)
    Next
    A = JnCrLf(O)
    Return
X_R:
    R = AyAlignL(VarAySqlQuote(Vy))
    Dim J1%
    For J1 = 0 To UB(R) - 1
        R(J1) = R(J1) + ","
    Next
    Return
End Function


Sub SpecCrtTbl()
DbCrtSpecTbl CurDb
End Sub

Sub SpecEnsTbl()
DbEnsSpecTbl CurDb
End Sub

Sub SpecExp()
SpecPthClr
Dim X
For Each X In AyNz(SpecNy)
    SpnmExp X
Next
End Sub

Function SpecNy() As String()
SpecNy = DbSpecNy(CurDb)
End Function

Function SpecPth$()
SpecPth = PthEns(CurDbPth & "Spec\")
End Function

Sub SpecPthBrw()
PthBrw SpecPth
End Sub

Sub SpecPthClr()
PthClr SpecPth
End Sub

Function SpecSchmy() As String()
SpecSchmy = SplitCrLf(SpecSchmLines)
End Function

Sub Stp()
Stop
End Sub

Function T0F2LinHasTF(A, T$, F$) As Boolean
Dim TLik$, FLikSsl$
LinAsgTRst A, TLik, FLikSsl
If T Like TLik Then
    If StrLikss(F, FLikSsl) Then
        T0F2LinHasTF = True
        Exit Function
    End If
End If
End Function


Property Get TFDes$(T$, F$)
TFDes = DbtfDes(CurDb, T, F)
End Property

Property Let TFDes(T$, F$, V$)
DbtfDes(CurDb, T, F) = V
End Property

Function TFLinHasPk(A$) As Boolean
TFLinHasPk = HasSubStr(A, " * ")
End Function

Function TFLinHasSk(A$) As Boolean
TFLinHasSk = HasSubStr(A, " | ")
End Function

Function TFTyChkMsg$(T, F, Ty As DAO.DataTypeEnum, ExpTyAy() As DAO.DataTypeEnum)
'DbtfTyMsg = FmtQQ("Table[?] field[?] has type[?].  It should be type[?].", T, F, S1, S2)

End Function

Function TimSz_XTSz$(A As Date, Sz&)
TimSz_XTSz = DteDTim(A) & "." & Sz
End Function

Function TkIsExist(T, K&) As Boolean
TkIsExist = DbtkIsExist(CurDb, T, K)
End Function

Sub TTCls(TT$)
AyDo CvNy(TT), "TblCls"
End Sub

Function UIxAy(U&) As Long()
Dim O&(), J&
ReDim O(U)
For J = 0 To U
    O(J) = J
Next
UIxAy = O
End Function

Function UniqFny() As String()
Stop '
'Dim I, M As LABC, O$()
'If IsEmp Then Exit Property
'For Each I In A
'    Set M = I
'    PushNoDupAy O, M.Fny
'Next
'UniqFny = O
End Function

Function XFyX$(A$(), F)
Dim L
For Each L In AyNz(A)
    XFyX = XFLinX(L, F)
    If XFyX <> "" Then Exit Function
Next
End Function

Function XFLinX$(A, F)
Dim X$, FLikss$
LinAsgTRst A, X, FLikss
If StrLikss(F, FLikss) Then XFLinX = X
End Function

Function XSqpInBExpr$(Ay, FldNm$, Optional WithQuote As Boolean)
Const C$ = "[?] in (?)"
Dim B$
    If WithQuote Then
        B = JnComma(AyQuoteSng(Ay))
    Else
        B = JnComma(Ay)
    End If
XSqpInBExpr = FmtQQ(C, FldNm, B)
End Function

Private Sub Z_ApScl()
Act = ApScl(" ", "")
Ept = ""
C
End Sub

Private Sub Z_DbAddTd()
Dim A As DAO.TableDef
TblDrp "Tmp"
Set A = DbAddTd(CurDb, TmpTd)
TblDrp "Tmp"
End Sub

Private Sub Z_DbtfAddExpr()
TblDrp "Tmp"
Dim A As DAO.TableDef
Set A = DbAddTd(CurDb, TmpTd)
DbtfAddExpr CurDb, "Tmp", "F2", "[F1]+"" hello!"""
TblDrp "Tmp"
End Sub

Private Sub Z_DrsEnsRplDbt()
Dim Db As Database, D1 As Drs, D2 As Drs
Set Db = TmpDb
Set D1 = SampleDrs
DrsEnsRplDbt D1, Db, "T"
Set D2 = DbtDrs(Db, "T")
Ass IsEqAy(D1, D2)
DbKill Db
End Sub

Private Sub Z_DrsInsUpdDbt()
Dim Db As Database, T$, A As Drs, TFb$
    TFb = TmpFb("Tst", "DrsInsUpdDbt")
    Set Db = FbCrt(TFb)
T = "Tmp"
Db.Execute "Create Table Tmp (A Int, B Int, C Int)"
Db.Execute CrtSkSql("Tmp", "A")
'DbSqyRun Db, InsDrApSqy("Tmp", "A B C", Array(1, 3, 4), Array(3, 4, 5))
Set A = Drs("A B C", CvAy(Array(Array(1, 2, 3), Array(2, 3, 4))))

Ept = Array(Array(1&, 2&, 3&), Array(2&, 3&, 4&), Array(3&, 4&, 5&))
GoSub Tst
Db.Close
Kill TFb
Exit Sub
Tst:
    DrsInsUpdDbt A, Db, T
    Act = DbtDry(Db, T)
    C
    Return
End Sub


Private Sub Z_LblSeqAy()
Dim Act$(), A, N%, Exp$()
A = "Lbl"
N = 10
Exp = SslSy("Lbl01 Lbl02 Lbl03 Lbl04 Lbl05 Lbl06 Lbl07 Lbl08 Lbl09 Lbl10")
Act = LblSeqAy(A, N)
Ass IsEqAy(Act, Exp)
End Sub

Private Sub Z_PfxSsl_Sy()
Dim A$, Exp$()
A = "A B C D"
Exp = SslSy("AB AC AD")
GoSub Tst
Exit Sub
Tst:
Dim Act$()
Act = PfxSsl_Sy(A)
Debug.Assert IsEqAy(Act, Exp)
Return
End Sub


Private Sub Z_SclChk()
Dim A$, Ny0
A = "Req;Alw;Sz=1"
Ny0 = VdtEleSclNmSsl
Ept = EmpSy
Push Ept, "There are [invalid-SclNy] in given [scl] under these [valid-SclNy]."
Push Ept, "    [invalid-SclNy] : Alw"
Push Ept, "    [scl]           : Req;Alw;Sz=1"
Push Ept, "    [valid-SclNy]   : Req AlwZLen Sz Dft VRul VTxt Des Expr"
GoSub Tst
Exit Sub
Tst:
    Act = SclChk(A, Ny0)
    C
End Sub

Private Sub ZZ_ApDtAy()
Dim A() As Dt
A = ApDtAy(SampleDt1, SampleDt2)
Stop
End Sub

Private Sub ZZ_DicHasStrKy()
ZZ_DicHasStrKy__X "DicHasStrKy"
End Sub

Private Sub ZZ_DicHasStrKy__X(X$)
Dim A As New Dictionary, Exp As Boolean, Act As Boolean
Dim J&
For J = 1 To 10000
    A.Add CStr(J), J
Next
Act = Run(X, A)
Exp = True
Ass Act = Exp

A.Add 10001, "X"
Act = Run(X, A)
Exp = False
Ass Act = Exp

End Sub

Private Sub ZZ_DicHasStrKy1()
ZZ_DicHasStrKy__X "DicHasStrKy1"
End Sub

Private Sub ZZ_DicHasStrKy2()
Dim A As New Dictionary, Exp As Boolean, Act As Boolean
Dim J&
For J = 1 To 10000
    A.Add CStr(J), J
Next
Act = DicHasStrKy(A)
Exp = True
Ass Act = Exp

A.Add 10001, "X"
Act = DicHasStrKy(A)
Exp = False
Ass Act = Exp

End Sub

Sub ZZ_ErAyzFxWsMissingCol()
'" [Material]             As Sku," & _
'" [Plant]                As Whs," & _
'" [Storage Location]     As Loc," & _
'" [Batch]                As BchNo," & _
'" [Unrestricted]         As OH " & _

End Sub

Sub ZZ_ReSeqSpecFny()
AyBrw ReSeqSpecFny("Flg RecTy Amt Key Uom MovTy Qty BchRateUX RateTy Bch Las GL |" & _
" Flg IsAlert IsWithSku |" & _
" Key Sku PstMth PstDte |" & _
" Bch BchNo BchPermitDate BchPermit |" & _
" Las LasBchNo LasPermitDate LasPermit |" & _
" GL GLDocNo GLDocDte GLAsg GLDocTy GLLin GLPstKy GLPc GLAc GLBusA GLRef |" & _
" Uom Des StkUom Ac_U")
End Sub

Private Sub ZZ_SslSqBktCsv()
Debug.Print SslSqBktCsv("a b c")
End Sub

Private Sub ZZ_TFDes()
TFDes("Att", "AttNm") = "AttNm"
End Sub

Private Function ZZCrdTyLvs$()
ZZCrdTyLvs = "1 2 3"
End Function

Function LyExt(NoT1$()) As String()
LyExt = LyXXX(NoT1, "Ext")
End Function

Function LyFld(NoT1$()) As String()
LyFld = LyXXX(NoT1, "Fld")
End Function


Private Function LyXXX(NoT1$(), XXX$) As String()
LyXXX = AyWhRmvT1(NoT1, XXX)
End Function


Private Function LyStuInp(NoT1$()) As String()
LyStuInp = LyXXX(NoT1, "StuInp")
End Function

