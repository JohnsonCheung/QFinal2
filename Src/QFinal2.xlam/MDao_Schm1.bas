Attribute VB_Name = "MDao_Schm1"
Option Explicit
Private Type T: Lno As Integer: T As String: Fny() As String: Sk() As String:     End Type
Private Type F: Lno As Integer: E As String: LikT As String:  LikFny() As String: End Type
Private Type DD: Lno As Integer: T As String: F As String:     Des As String:      End Type
Private Type E
    Lno As Integer
    E As String
    Ty As DAO.DataTypeEnum
    Req As Boolean
    ZLen As Boolean
    TxtSz As Byte
    Expr As String
    VRul As String
    Dft As String
    VTxt As String
End Type

Sub DbSchmCrt1(A As Database, SchmLines$)
Dim Er$(), Td() As DAO.TableDef, Pk$(), Sk$(), TDes$(), FDes$()
GoRslt SchmLines, Er, Td, Pk, Sk, TDes, FDes
AyBrwThw Er
AyDoPX Td, "DbAddTd", A
AyDoPX Pk, "DbRun", A
AyDoPX Sk, "DbRun", A
AyDoPX TDes, "DbSetTDes", A
AyDoPX FDes, "DbSetFDes", A
End Sub

Private Sub GoRslt(SchmLines$, OEr$(), OTd() As DAO.TableDef, OPk$(), OSk$(), OTDes$(), OFDes$())
Dim Er$(), T() As T, F() As F, E() As E, D() As DD, Tny$(), Eny$()
GoBrk SchmLines, Er, T, F, E, D, Tny, Eny
OEr = MkEr(Er, T, F, E, D, Tny, Eny): If Sz(OEr) > 0 Then Exit Sub
OTd = MkTd(Tny, T, F, E)
OPk = MkPk(Tny, T)
OSk = MkSk(Tny, T)
OFDes = MkFDes(Tny, T, D)
OTDes = MkTDes(Tny, D)
End Sub

Private Function MkEr(Er$(), T() As T, F() As F, E() As E, D() As DD, Tny$(), Eny$()) As String()
MkEr = AyAddAp _
    (Er _
   , Er_DupT(T, Tny) _
   , Er_DupE(E, Eny) _
   , Er_TzDLy_NotIn_Tny(D, Tny) _
   , Er_FzDLy_NotIn_TblFny(D, Tny, T) _
   , Er_EzFLy_NotIn_Eny(F, Eny) _
    )
End Function

Private Function MkFDes(Tny$(), T() As T, D() As DD) As String()
Stop '
End Function

Private Function MkPk(Tny$(), T() As T) As String()
Dim Tbl
For Each Tbl In Tny
    PushNonEmp MkPk, FndPkSql(Tbl, T)
Next
End Function

Private Function MkSk(Tny$(), T() As T) As String()
Dim Tbl
For Each Tbl In Tny
    PushNonEmp MkSk, FndSkSql(Tbl, T)
Next
End Function

Private Function MkTd(Tny$(), T() As T, F() As F, E() As E) As DAO.TableDef()
Dim Tbl
For Each Tbl In Tny
    PushObj MkTd, NewTd(Tbl, FndFdAy(Tbl, Tny, T, F, E))
Next
End Function

Private Function MkTDes(Tny$(), D() As DD) As String()
Stop '
End Function

Private Sub GoBrk(SchmLines$, OEr$(), OT() As T, OF_() As F, OE() As E, OD() As DD, OTny$(), OEny$())
Dim Er$()
Dim ClnLnxAy()  As Lnx
Dim E()  As Lnx
Dim F()  As Lnx
Dim D()  As Lnx
Dim T()  As Lnx
ClnLnxAy = LyClnLnxAy(SplitCrLf(SchmLines))
T = LnxAyWhRmvT1(ClnLnxAy, "T")
D = LnxAyWhRmvT1(ClnLnxAy, "D")
E = LnxAyWhRmvT1(ClnLnxAy, "E")
F = LnxAyWhRmvT1(ClnLnxAy, "F")

Dim TEr$(), FEr$(), EEr$(), DEr$()
OE = BrkE(E, EEr)
OF_ = BrkF(F, FEr)
OD = BrkD(D, DEr)
OT = BrkT(T, TEr)
Er = LnxAyT1Chk(ClnLnxAy, "D E F T")
OEr = CvSy(AyAddAp(Er, TEr, FEr, EEr, DEr))
OEny = AyTakT1(LnxAyLy(E))
OTny = AyTakT1(LnxAyLy(T))
End Sub

Private Function BrkT(A() As Lnx, OEr$()) As T()
If Sz(A) = 0 Then
    Push OEr, ErMsg_NoTLin
    Exit Function
End If
Dim J%
For J = 0 To UB(A)
    XPushT BrkT, BrkTLin(A(J), OEr)
Next
End Function

Private Function BrkTLin(T As Lnx, OEr$()) As T
If Not HasSubStr(T.Lin, "|") Then
    Push OEr, "should have a |"
    Exit Function
End If
Dim A$, B$, C$, D$
BrkAsg T.Lin, "|", A, B
With BrkTLin
    .T = A
    B = Replace(B, "*", A)
    BrkS1Asg B, "|", C, D
    If D = "" Then
        .Fny = SslSy(C)
    Else
        .Sk = SslSy(RmvPfx(C, A & " "))
        .Fny = SslSy(Replace(B, "|", " "))
    End If
    If Sz(.Fny) = 0 Then
        Push OEr, "should have fields after |"
        Exit Function
    End If
    Dim Dup$()
    Dup = AyWhDup(.Fny)
    If Sz(Dup) > 0 Then
        Stop '
'       Push BrkTLin.Er, ErMsg_DupF(T.Ix + 1)
        Exit Function
    End If
End With
End Function


Private Sub Z_BrkTLin()
Dim Act As T
Dim Ept As T
Dim Emp As T
Dim EptEr$()
Dim TLin As Lnx
Set TLin = Lnx(999, "A")
Ept = Emp
Push EptEr, "should have a |"
GoSub Tst
'
TLin.Lin = "A | B B"
Ept = Emp
Push EptEr, "dup fields[B]"
GoSub Tst
'
TLin.Lin = "A | B B D C C"
Ept = Emp
Push EptEr, "dup fields[B C]"
GoSub Tst
'
TLin.Lin = "A | * B D C"
Ept = Emp
With Ept
    .T = "A"
    .Fny = SslSy("A B D C")
End With
GoSub Tst
'
TLin = "A | * B | D C"
Ept = Emp
With Ept
    .T = "A"
    .Fny = SslSy("A B D C")
    .Sk = SslSy("B")
End With
GoSub Tst
'
TLin = "A |"
Ept = Emp
Push EptEr, "should have fields after |"
GoSub Tst
Exit Sub
Tst:
    Dim ActEr$()
    Act = BrkTLin(TLin, ActEr)
    Ass IsEqAy(ActEr, EptEr)
    Ass ZZ_IsTItmEq(Act, Ept)
    Return
End Sub

Private Function BrkD(D() As Lnx, OEr$()) As DD()
Dim J%
For J = 0 To UB(D)
    XPushD BrkD, BrkDLin(D(J), OEr)
Next
End Function

Private Function BrkDLin(D As Lnx, OEr$()) As DD
Dim V$
With BrkDLin
    AyAsg Lin3TRst(D.Lin), .T, .F, V, .Des
    If V <> "|" Then Push OEr, "..."
End With
End Function

Private Function BrkE(A() As Lnx, OEr$()) As E()
Dim J%
For J = 0 To UB(A)
    XPushE BrkE, BrkELin(A(J), OEr)
Next
End Function

Private Function BrkELin(ELin As Lnx, OEr$()) As E
Dim LikFF1$, V$, Ty$, Ay(), Brk$(), Rest$(), Itm$()
Itm = LinTermAy(ELin.Lin)
With BrkELin
    AyAsg ShfLblVal(Itm, "*Ty ?Req ?ZLen TxtSz VTxt Dft VRul Expr"), _
                           Ty, .Req, .ZLen, .TxtSz, .VTxt, .Dft, .VRul, .Expr
    .Ty = DaoShtTyStrTy(Ty)
    If Sz(Itm) > 0 Then
        Push OEr, ErMsg_ExcessEleItm(ELin.Ix, JnSpc(Itm))
    End If
    If Ty = 0 Then
        Push OEr, ErMsg_TyEr(ELin.Ix, Ty)
    End If
End With
End Function


Private Function BrkF(A() As Lnx, OEr$()) As F()
Dim J%
For J = 0 To UB(A)
    XPushF BrkF, BrkFLin(A(J), OEr)
Next
End Function

Private Function BrkFLin(F As Lnx, OEr$()) As F
Dim LikFF$, A$, V$
With BrkFLin
    AyAsg Lin3TRst(F.Lin), .E, .LikT, V, A
    .LikFny = SslSy(LikFF)
End With
End Function


Private Function Er_DupE(A() As E, Eny$()) As String()
Dim Dup$(), IE, E$, LnoAy%()
Dup = AyWhDup(Eny)
If Sz(Dup) = 0 Then Exit Function
For Each IE In Dup
    E = IE
    LnoAy = FndELnoAy(E, A)
    Push Er_DupE, ErMsg_DupE(LnoAy, E)
Next
End Function

Private Function Er_DupT(A() As T, Tny$()) As String()
Dim Dup$(), IT, T$, LnoAy%()
Dup = AyWhDup(Tny)
If Sz(Dup) = 0 Then Exit Function
For Each IT In Dup
    T = IT
    LnoAy = FndTLnoAy(T, A)
    Push Er_DupT, ErMsg_DupT(LnoAy, T)
Next
End Function

Private Function Er_EzFLy_NotIn_Eny(F() As F, Eny$()) As String()
Dim J%, O$()
For J = 0 To XFUB(F)
    With F(J)
        Stop '
        'If Not AyHas(Eny, .E) Then Push O, ErMsg_EzFLy_NotIn_Eny(.Lno, .E)
    End With
Next
Er_EzFLy_NotIn_Eny = O
End Function

Private Function Er_FzDLy_NotIn_TblFny(D() As DD, Tny$(), T() As T) As String()
Dim J%, Fny1$()
For J = 0 To XDUB(D)
    With D(J)
        If Not AyHas(Tny, .T) Then GoTo Nxt
        Fny1 = FndFny(.T, T)
        If Not AyHas(Fny1, .F) Then
            Push Er_FzDLy_NotIn_TblFny, ErMsg_FzDLy_NotIn_TblFny(.Lno, .T, .F, JnSpc(Fny1))
        End If
    End With
Nxt:
Next
End Function

Private Function Er_TzDLy_NotIn_Tny(D() As DD, Tny$()) As String()
Dim Tssl$, J%
Tssl = JnSpc(Tny)
For J = 0 To XDUB(D)
    With D(J)
        If Not AyHas(Tny, .T) Then
            Push Er_TzDLy_NotIn_Tny, ErMsg_TzDLy_NotIn_Tny(.Lno, .T, Tssl)
        End If
    End With
Next
End Function

Private Function FndPkSql$(Tbl, T() As T)
With FndT(Tbl, T)
    If Not AyHas(.Fny, .T) Then Exit Function
End With
FndPkSql = CrtPkSql(Tbl)
End Function

Private Function FndSkSql$(Tbl, T() As T)
FndSkSql = CrtSkSql(Tbl, FndT(Tbl, T).Sk)
End Function

Private Function FndE(Tbl, Fld, F() As F, E() As E) As E
Dim J%, O As F, M As F
For J = 0 To UBound(F)
    M = F(J)
    If Tbl Like M.LikT Then
        If StrLikssAy(Fld, M.LikFny) Then
            FndE = FndE__1(M.E, E)
            If FndE.E <> M.E Then Stop
            Exit Function
        End If
    End If
Next
End Function

Private Function FndE__1(Ele$, E() As E) As E
Dim J%
For J = 0 To UBound(E)
    If E(J).E = Ele Then
        FndE__1 = E(J)
        Exit Function
    End If
Next
End Function

Private Function FndELnoAy(E$, EBrk() As E) As Integer()
Dim J%
For J = 0 To UBound(EBrk)
    Push FndELnoAy, EBrk(J).Lno
Next
End Function

Private Function FndFd(Tbl, Fld, Tny$(), F() As F, E() As E) As DAO.Field2
Select Case True
Case Tbl = Fld:       Set FndFd = NewFd_zId(Fld)
Case AyHas(Tny, Tbl): Set FndFd = NewFd_zFk(Fld)
Case Else
With FndE(Tbl, Fld, F, E)
    Set FndFd = NewFd(Fld, .Ty, .TxtSz, .ZLen, .Expr, .Dft, .Req, .VRul, .VTxt)
End With
End Select
End Function

Private Function FndFdAy(Tbl, Tny$(), T() As T, F() As F, E() As E) As DAO.Field2()
Dim Fld, O() As DAO.Field2
For Each Fld In FndFny(Tbl, T)
    PushObj O, FndFd(Tbl, Fld, Tny, F, E)
Next
FndFdAy = O
End Function

Private Function FndFny(Tbl, T() As T) As String()
Dim J%
With FndT(Tbl, T)
    FndFny = .Fny
    If .T <> Tbl Then Stop
End With
End Function


Private Function FndT(Tbl, T() As T) As T
Dim J%
For J = 0 To UBound(T)
    With T(J)
        If .T = Tbl Then FndT = T(J): Exit Function
    End With
Next
End Function

Private Function FndTLnoAy(Tbl, T() As T) As Integer()
Dim J%
For J = 0 To XTUB(T)
    PushI FndTLnoAy, T(J).Lno
Next
End Function

Private Function ErMsg$(Lno%, M$)
ErMsg = "--Lno" & Lno & ".  " & M
End Function

Private Function ErMsg_DupE$(LnoAy%(), E$)
ErMsg_DupE = ErMsg1(LnoAy, FmtQQ("This E[?] is dup", E))
End Function

Private Function ErMsg_DupF$(Lno%, T$, Fny$())
ErMsg_DupF = ErMsg(Lno, FmtQQ("F[?] is dup in T[?]", JnSpc(Fny), T))
End Function

Private Function ErMsg_DupT$(LnoAy%(), T$)
ErMsg_DupT = ErMsg1(LnoAy, FmtQQ("This T[?] is dup", T))
End Function

Private Function ErMsg_ExcessEleItm$(Lno%, L$)
ErMsg_ExcessEleItm = ErMsg(Lno, FmtQQ("Excess Ele Item[?]", L))
End Function

Private Function ErMsg_ExcessTxXTSz$(Lno%, Ty$)
ErMsg_ExcessTxXTSz = ErMsg(Lno, FmtQQ("Ty[?] is not Txt, it should not have TxtSz", Ty))
End Function


Private Function ErMsg_EzFLy_NotIn_Eny$(Lno%, E$, Essl$)
ErMsg_EzFLy_NotIn_Eny = ErMsg(Lno, FmtQQ("E[?] of is not in E-Lin[?]", E, Essl))
End Function

Private Function ErMsg_FldEleEr$(Lno%, E$, Essl$)
ErMsg_FldEleEr = ErMsg(Lno, FmtQQ("E[?] is invalid.  Valid E is [?]", E, Essl))
End Function

Private Function ErMsg_FzDLy_NotIn_TblFny$(Lno%, T$, F$, Fssl$)
ErMsg_FzDLy_NotIn_TblFny = ErMsg(Lno, FmtQQ("F[?] is invalid in T[?].  Valid F[?]", F, T, Fssl))
End Function

Private Function ErMsg_NoELin$()
ErMsg_NoELin = "No E-Line"
End Function


Private Function ErMsg_NoFLin$()
ErMsg_NoFLin = "No F-Line"
End Function

Private Function ErMsg_NoTLin$()
ErMsg_NoTLin = "No T-Line"
End Function

Private Function ErMsg_TblFldEr$(Lno%, T$, F$)
ErMsg_TblFldEr = ErMsg(Lno, FmtQQ("T[?] has invalid F[?], which cannot be found in any F-Lines"))
End Function

Private Function ErMsg_TyEr$(Lno%, Ty$)
ErMsg_TyEr = ErMsg(Lno, FmtQQ("Invalid DaoShtTy[?].  Valid ShtTy[?]", Ty, DaoShtTySsl))
End Function

Private Function ErMsg_TzDLy_NotIn_Tny$(Lno%, T$, Tssl$)
ErMsg_TzDLy_NotIn_Tny = ErMsg(Lno, FmtQQ("T[?] is invalid.  Valid T[?]", T, Tssl))
End Function

Private Function ErMsg1(LnoAy%(), M$)
ErMsg1 = "--" & Join(AyAddPfx(LnoAy, "Lno"), ".") & "  " & M
End Function

Private Sub XPushD(O() As DD, M As DD): Dim N&: N = XDSz(O): ReDim Preserve O(N): O(N) = M: End Sub
Private Sub XPushE(O() As E, M As E): Dim N&: N = XESz(O): ReDim Preserve O(N): O(N) = M: End Sub
Private Sub XPushF(O() As F, M As F): Dim N&: N = XFSz(O): ReDim Preserve O(N): O(N) = M: End Sub
Private Sub XPushT(O() As T, M As T): Dim N&: N = XTSz(O): ReDim Preserve O(N): O(N) = M: End Sub
Private Function XFSz%(A() As F): On Error Resume Next: XFSz = UBound(A) + 1: End Function
Private Function XTSz%(A() As T): On Error Resume Next: XTSz = UBound(A) + 1: End Function
Private Function XDSz%(A() As DD): On Error Resume Next: XDSz = UBound(A) + 1: End Function
Private Function XESz%(A() As E): On Error Resume Next: XESz = UBound(A) + 1: End Function
Private Function XFUB%(A() As F): XFUB = XFSz(A) - 1: End Function
Private Function XDUB%(A() As DD): XDUB = XDSz(A) - 1: End Function
Private Function XEUB%(A() As E): XEUB = XESz(A) - 1: End Function
Private Function XTUB%(A() As T): XTUB = XTSz(A) - 1: End Function


Private Function ZZ_IsTItmEq(A As T, B As T) As Boolean
If A.T <> B.T Then Exit Function
If Not IsEqAy(A.Fny, B.Fny) Then Exit Function
ZZ_IsTItmEq = IsEqAy(A.Sk, B.Sk)
End Function

