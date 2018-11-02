Attribute VB_Name = "MIde_Z_Src"
Option Explicit
Sub Z()
Z_SrcDic
Z_SrcMthDDNy
Z_SrcMthNy
End Sub
Function CurSrc() As String()
CurSrc = MdSrc(CurMd)
End Function

Private Sub ZZ_SrcContLin()
Dim O$(3)
O(0) = "A _"
O(1) = "  B _"
O(2) = "C"
O(3) = "D"
Dim Act$: Act = SrcContLin(O, 0)
Ass Act = "A B C"
End Sub

Private Sub ZZ_SrcDcl()
StrBrw SrcDclLy(ZZSrc)
End Sub

Private Sub ZZ_SrcDclLinCnt()
Dim B1$(): B1 = CurSrc
Dim B2$(): B2 = SrcSrtLy(B1)
Dim A1%: A1 = SrcDclLinCnt(B1)
Dim A2%: A2 = SrcDclLinCnt(SrcSrtLy(B1))
End Sub

Private Sub ZZ_SrcFstMthIx()
Dim Act%
Act = SrcFstMthIx(ZZSrc)
Ass Act = 2
End Sub

Private Sub ZZ_SrcMthBdyLy()
Dim Src$(): Src = ZZSrc
Dim MthNm$: MthNm = "A"
Dim Act$()
Act = SrcMthBdyLy(Src, MthNm)
End Sub

Private Sub ZZ_SrcMthIxTopRmkFm()
Dim ODry()
    Dim Src$(): Src = MdSrc(Md("IdeSrcLin"))
    Dim Dr(), Lx&
    Dim J%, IsMth$, RmkLx$, L
    For Each L In Src
        IsMth = ""
        RmkLx = ""
        If LinIsMth(CStr(L)) Then
            If Lx = 482 Then Stop
            IsMth = "*Mth"
            RmkLx = SrcMthIxTopRmkFm(Src, Lx)

        End If
        Dr = Array(IsMth, RmkLx, L)
        Push ODry, Dr
        Lx = Lx + 1
    Next
DrsBrw Drs("Mth RmkLx Lin", ODry)
End Sub

Private Sub ZZ_SrcMthLxAy()
Dim Src$(): Src = ZZSrc
Dim LxAy&(): LxAy = SrcMthLxAy(ZZSrc)
Dim Ay$(): Ay = AyWhIxAy(Src, LxAy)
Dim Dry(): Dry = AyZip(LxAy, Ay)
Dim O$()
O = DrsFmt(Drs("Lx Lin", AyZip(LxAy, Ay)))
PushAy O, DrsFmt(AyDrs(Src))
AyBrw O
End Sub

Private Sub ZZ_SrcMthLxAy1()
Dim Src$(): Src = MdSrc(Md("DaoDb"))
Dim Ay$(): Ay = AyWhIxAy(Src, SrcMthLxAy(Src))
AyBrw Ay
End Sub

Private Sub ZZ_SrcMthNy()
Dim Act$()
   Act = SrcMthNy(ZZSrc)
   AyBrw Act
End Sub

Private Function ZZSrc() As String()
ZZSrc = MdSrc(Md("IdeSrc"))
End Function

Private Function ZZSrcLin$()
ZZSrcLin = "Private Sub LinIsMth()"
End Function
Private Sub Z_SrcMthNy()
Brw SrcMthNy(MdSrc(Md("AAAMod")))
End Sub

Function SrcMthNy(A$(), Optional B As WhMth) As String()
SrcMthNy = AyWhDist(AyTakBefDot(MthDDNyWh(SrcMthDDNy(A), B)))
End Function

Function SrcDicTopRmkMthLinesAy(SrcDic As Dictionary) As String()
Dim L
For Each L In SrcDic.Items
    If FstChr(L) = "'" Then
        PushI SrcDicTopRmkMthLinesAy, L
    End If
Next
End Function

Function SrcMthKeyLinesDic1(A$()) As Dictionary
'To be delete
'Dim Ix, O As New Dictionary
'SrcDicAddDcl O, A
'For Each Ix In AyNz(SrcMthIx(A))
'    O.Add LinMthKey(A(Ix)), SrcMthIxLinesWithTopRmk(A, Ix)
'Next
'Set SrcMthKeyLinesDic = O
End Function

Function Src(MdNm$) As String()
Src = MdLy(Md(MdNm))
End Function

Function SrcAddMthIfNotExist(A$(), MthNm$, NewMthLy$()) As String()
If SrcHasMth(A, MthNm) Then
   SrcAddMthIfNotExist = A
Else
   SrcAddMthIfNotExist = AyAddAp(A, NewMthLy)
End If
End Function

Function SrcBdyLines$(A$())
SrcBdyLines = JnCrLf(SrcBdyLy(A))
End Function

Function SrcBdyLnoCnt(A$()) As LnoCnt
Dim Lno&
Dim Cnt&
   Lno = SrcDclLinCnt(A) + 1
   Cnt = Sz(A) - Lno + 1
Set SrcBdyLnoCnt = LnoCnt(Lno, Cnt)
End Function

Function SrcBdyLy(A$()) As String()
SrcBdyLy = AyWhFm(A, SrcDclLinCnt(A))
End Function

Function SrcCmpFmt(A1$(), A2$()) As String()
Dim D1 As Dictionary: Set D1 = SrcDic(A1)
Dim D2 As Dictionary: Set D2 = SrcDic(A2)
SrcCmpFmt = DicCmpFmt(D1, D2)
End Function

Function SrcContLin$(A$(), Ix)
If Ix <= -1 Then Exit Function
Const CSub$ = "SrcContLinFm"
Dim J&, I$
Dim O$, IsCont As Boolean
For J = Ix To UB(A)
   I = A(J)
   O = O & LTrim(I)
   IsCont = HasSfx(O, " _")
   If IsCont Then O = RmvSfx(O, " _")
   If Not IsCont Then Exit For
Next
If IsCont Then Er CSub, "each lines {Src} ends with sfx _, which is impossible"
SrcContLin = O
End Function


Function SrcDisMthNy(A$()) As String()
Dim O$(), I
If Sz(A) = 0 Then Exit Function
For Each I In A
   PushNonEmp O, LinMthNm(CStr(I))
Next
SrcDisMthNy = O
End Function

Function SrcEnsMth(T$(), MthNm$, NewMthLy$()) As String()
SrcEnsMth = SrcAddMthIfNotExist(T, MthNm, NewMthLy)
End Function

Function SrcHasMth(A$(), MthNm) As Boolean
SrcHasMth = SrcMthNmIx(A, MthNm) >= 0
End Function
Function SrcLinInfFny() As String()
Static X As Boolean, Y$()
If Not X Then
    X = True
    Y = SplitSpc("Md Lno Lin EnmNm IsBlank IsEmn IsMth LinIsPrp IsRmk IsTy Mdy MthNm MthTy NoMdy PrpTy TyNm")
End If
SrcLinInfFny = Y
End Function

Private Function SrcLxAy(A$()) As Long()
Dim L$, Lx%, J%, MthNm$
Dim O1&(): O1 = SrcMthLxAy(A)
Dim O&()
    For J = 0 To UB(O1)
        Lx = O1(J)
        L = A(Lx)
        If HasPfx(L, "Private ") Then GoTo Nxt
        MthNm = LinMthNm(L)
        If Not HasPfx(MthNm, "ZZ") Then GoTo Nxt
        Push O, O1(J)
Nxt:
    Next
SrcLxAy = O
End Function

Function SrcMthBdyLines$(A$(), MthNm)
SrcMthBdyLines = JnCrLf(SrcMthBdyLy(A, MthNm))
End Function

Function SrcMthBdyLy(A$(), MthNm) As String()
Dim FTIx() As FTIx: FTIx = SrcMthFTIxAy(A, MthNm)
Dim O$(), J%
For J = 0 To UB(FTIx)
   PushAy O, AyWhFTIx(A, FTIx(J))
Next
SrcMthBdyLy = O
End Function

Sub Z_SrcMthBdyLy()
Dim MthNm$: MthNm = "A"
Dim Act$()
Act = SrcMthBdyLy(CurSrc, MthNm)
End Sub

Function SrcMthLnoCntAy(A$(), MthNm) As LnoCnt()
Dim FmAy&(): FmAy = SrcMthNmIxAy(A, MthNm)
Dim O() As LnoCnt, J%
Dim ToIx&
Dim FTIx As FTIx
Dim LnoCnt As LnoCnt
For J = 0 To UB(FmAy)
   ToIx = SrcMthLx_ToLx(A, FmAy(J))
   FTIx = FTIx(FmAy(J), ToIx)
   LnoCnt = FTIxLnoCnt(FTIx)
   PushObj O, LnoCnt
Next
SrcMthLnoCntAy = O
End Function


Sub SrcMthDrAsg(A, OShtMdy$, OShtTy$, ONm$, OPrm$, ORet$, OLinRmk$, OLines$, OTopRmk$)
AyAsg A, OShtMdy, OShtTy, ONm, OPrm, ORet, OLinRmk, OLines, OTopRmk
End Sub

Function SrcMthFT(A$()) As FTIx()
Dim F&(): F = SrcMthIx(A)
Dim U%: U = UB(F)
If U = -1 Then Exit Function
Dim O() As FTIx
ReDim O(U)
Dim J%
For J = 0 To U
    Set O(J) = FTIx(F(J), SrcMthIxTo(A, F(J)))
Next
SrcMthFT = O
End Function
Function SrcMthFTAy(A$(), MthNm) As FTIx()
Dim F&()
F = SrcMthNmIxAy(A, MthNm): If Sz(F) <= 0 Then Exit Function
Dim O() As FTIx
ReDim O(UB(F))
Dim J%
For J = 0 To UB(F)
    Set O(J) = FTIx(F(J), SrcMthIxTo(A, F(J)))
Next
SrcMthFTAy = O
End Function

Function SrcMthIxLinesWithTopRmk$(A$(), MthIx&)
Dim B$: B = SrcMthIxTopRmk(A, MthIx)
Dim C$: C = SrcMthIxLines(A, MthIx)
If B <> "" Then C = B & vbCrLf & C
SrcMthIxLinesWithTopRmk = C
End Function

Function SrcMthIxLines$(A$(), MthIx&, Optional WithTopRmk As Boolean)
Dim L2&, Fm&
L2 = SrcMthIxTo(A, MthIx): If L2 = 0 Then Stop
If WithTopRmk Then
    Fm = SrcMthIxTopRmkFm(A, MthIx)
Else
    Fm = MthIx
End If
SrcMthIxLines = Join(SyWhFmTo(A, Fm, L2), vbCrLf)
End Function

Sub SrcDicAddDcl(A As Dictionary, Src$())
Dim Dcl$
Dcl = SrcDclLines(Src)
If Dcl = "" Then Exit Sub
A.Add "*Dcl", Dcl
End Sub

Private Sub Z_SrcDic()
DicBrw SrcDic(MdSrc(Md("AAAMod")))
End Sub

Function SrcDic(A$()) As Dictionary
Set SrcDic = SrcMthDDNmDic(A)
End Function

Function SrcMthKy(A$()) As String()
Dim Ix
For Each Ix In AyNz(SrcMthIx(A))
    PushI SrcMthKy, LinMthSrtKey(A(Ix))
Next
End Function
Function SrcMthLinDryWP(A$()) As Variant()
Dim L
For Each L In AyNz(A)
    PushISomSz SrcMthLinDryWP, LinMthDrWP(L)
Next
End Function

Function SrcMthLines$(A$(), MthNm, Optional WithTopRmk As Boolean)
Dim I, O$()
For Each I In AyNz(SrcMthNmIxAy(A, MthNm))
    PushI O, SrcMthIxLines(A, CLng(I), WithTopRmk)
Next
SrcMthLines = Join(O, vbCrLf & vbCrLf)
End Function

Function SrcMthLinesDic(A$(), Optional ExlDcl As Boolean) As Dictionary
Dim L&(): L = SrcMthIx(A)
Dim O As New Dictionary
    If Not ExlDcl Then O.Add "*Dcl", SrcDclLines(A)
    If Sz(L) = 0 Then GoTo X
    Dim MthNm$, Lin$, Lines$, Lx
    For Each Lx In L
        Lin = A(Lx)
        MthNm = LinMthNm(Lin):            If MthNm = "" Then Stop
        Lines = SrcMthIxLines(A, CLng(Lx)): If Lines = "" Then Stop
        If O.Exists(MthNm) Then
            If Not LinIsPrp(Lin) Then Stop
            O(MthNm) = O(MthNm) & vbCrLf & vbCrLf & Lines
        Else
            O.Add MthNm, Lines
        End If
    Next
X:
Set SrcMthLinesDic = O
End Function

Function SrcMthLx_ToLx&(A$(), MthLx)
Const CSub$ = "SrcMthLx_ToLx"
Dim Lin$
   Lin = A(MthLx)

Dim Pfx$
'   Pfx = SrcLin_EndLinPfx(Lin)
Dim O&
   For O = MthLx + 1 To UB(A)
       If HasPfx(A(O), Pfx) Then SrcMthLx_ToLx = O: Exit Function
   Next
Er CSub, "{Src}-{MthFmIx} is {MthLin} which does have {FunEndLinPfx} in lines after [MthFmIx]", A, MthLx, Lin, Pfx
End Function

Function SrcMthLxAy(A$()) As Long()
If Sz(A) = 0 Then Exit Function
Dim O&(), I, J&
   For Each I In A
       If LinIsMth(CStr(I)) Then PushI O, J
       J = J + 1
   Next
SrcMthLxAy = O
End Function

Function SrcMthNmFC(A$(), MthNm) As FmCnt()
SrcMthNmFC = FTIxAyFC(SrcMthNmFT(A, MthNm))
End Function

Function SrcMthNmFT(A$(), MthNm) As FTIx()
Dim ToIx%, Fm
For Each Fm In AyNz(SrcMthNmIx(A, MthNm))
    ToIx = SrcMthIxTo(A, CLng(Fm))
    Push SrcMthNmFT, FTIx(Fm, ToIx)
Next
End Function

Private Sub Z_SrcMthDDNy()
Brw SrcMthDDNy(MdSrc(Md("AAAMod")))
End Sub

Function SrcMthDDNy(A$()) As String()
Dim L
For Each L In AyNz(A)
    PushNonBlankStr SrcMthDDNy, LinMthDDNm(CStr(L))
Next
End Function

Function SrcMthBrkAy(A$()) As Variant()
Dim L
For Each L In AyNz(A)
    PushNonZSz SrcMthBrkAy, LinMthNmBrk(CStr(L))
Next
End Function

Function SrcMthCnt%(A$(), Optional B As WhMth)
SrcMthCnt = Sz(SrcMthIx(A, B))
End Function

Function SrcNDisMth%(A$())
SrcNDisMth = Sz(SrcDisMthNy(A))
End Function

Function SrcNMth%(A$(), Optional B As WhMth)
SrcNMth = SrcMthCnt(A, B)
End Function

Function SrcNTy%(A$())
If Sz(A) = 0 Then Exit Function
Dim I, O%
For Each I In A
'   If SrcLin_IsTy(I) Then O = O + 1
Next
SrcNTy = O
End Function

Function SrcPth$()
Dim X As Boolean, Y$
If Not X Then
    X = True
    Y = CurDbPth & "Src\"
    PthEns Y
End If
SrcPth = Y
End Function

Sub SrcPthBrw()
PthBrw SrcPth
End Sub

Function SrcRmvMth(A$(), MthNm) As String()
Dim FTIxAy() As FTIx
   FTIxAy = SrcMthFTIxAy(A, MthNm)
Dim O$()
   O = A
   Dim J%
   For J = UB(FTIxAy) To 0 Step -1
       O = AyRmvFTIx(O, FTIxAy(J))
   Next
SrcRmvMth = O
End Function

Function SrcRmvTy(A$(), TyNm$) As String()
SrcRmvTy = AyRmvFTIx(A, DclTyFTIx(A, TyNm))
End Function

Function SrcRplMth(A$(), MthNm$, NewMthLy$()) As String()
Dim OldMthLines$
   OldMthLines = SrcMthBdyLines(A, MthNm)
Dim NewMthLines$
   NewMthLines = JnCrLf(NewMthLy)
If OldMthLines = NewMthLines Then
   SrcRplMth = A
   Exit Function
End If
Dim O$()
   O = SrcRmvMth(A, MthNm)
   PushAy O, NewMthLy
SrcRplMth = O

End Function

Function Srcy() As String()
Srcy = DbSrcTny(CurrentDb)
End Function
