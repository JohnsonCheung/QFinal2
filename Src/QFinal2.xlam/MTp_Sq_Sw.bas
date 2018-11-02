Attribute VB_Name = "MTp_Sq_Sw"
Option Explicit
Const CMod$ = "TpSqSw."
Private A_Pm As Dictionary
Sub Z()
Z_FndSwAsg
End Sub

Private Sub AAMain()
Z_FndSwAsg
End Sub

Private Function ChkDupNm(IO() As SwBrk) As String()
Dim Ny$(), Nm$, O() As SwBrk
Dim J%, M As SwBrk, Er() As SwBrk
For J = 0 To UB(IO)
    Set M = IO(J)
    If AyHas(Ny, M.Nm) Then
        PushObj Er, M
    Else
        PushObj O, M
        PushI Ny, M.Nm
    End If
Next
ChkDupNm = MsgDupNm(Er)
IO = O
End Function

Private Function ChkFld(A() As SwBrk, OEr$()) As SwBrk()
ChkFld = A
Exit Function
Dim M As SwBrk, IsEr As Boolean, J%, I, A1() As SwBrk, A2() As SwBrk
IsEr = True
A1 = A
While IsEr
    J = J + 1: If J > 1000 Then Stop
    IsEr = False
    For Each I In AyNz(A1)
        Set M = ChkFldLin(CvSwBrk(I), SwNmDic(A), OEr)
        If IsNothing(M) Then
            IsEr = True
        Else
            PushObj A2, M
        End If
    Next
    A1 = A2
Wend
ChkFld = A2
End Function

Private Function ChkFldLin(A As SwBrk, SwNm As Dictionary, OEr$()) As SwBrk
'Each Term in A.TermAy must be found either in Sw or Pm
Dim O0$(), O1$(), O2$(), I
For Each I In AyNz(A.TermAy)
    Select Case True
    Case HasPfx(I, "?"):  If Not SwNm.Exists(I) Then Push O0, I
    Case HasPfx(I, "@?"): If Not A_Pm.Exists(I) Then Push O1, I
    Case Else:                  Push O2, I
    End Select
Next
PushIAy OEr, MsgTermNotInSw(O0, A, SwNm)
PushIAy OEr, MsgTermNotInPm(O1, A)
PushIAy OEr, MsgTermMustBegWithQuestOrAt(O2, A)
If AyApHasEle(O0, O1, O2) Then Set ChkFldLin = A
End Function

Private Function ChkLin1$(IO As SwBrk)
Dim Msg$
With IO
    If .Nm = "" Then ChkLin1 = MsgNoNm(IO): Exit Function
    Select Case .OpStr
    Case "OR", "AND": If Sz(.TermAy) = 0 Then ChkLin1 = MsgTermCntAndOr(IO): Exit Function
    Case "EQ", "NE":  If Sz(.TermAy) <> 2 Then ChkLin1 = MsgTermCntEqNe(IO): Exit Function
    Case Else:        ChkLin1 = MsgOpStrEr(IO): Exit Function
    End Select
End With
End Function

Private Function ChkLin(IO() As SwBrk) As String()
Dim I
For Each I In AyNz(IO)
    PushNonNothing ChkLin, ChkLin1(CvSwBrk(I))
Next
End Function

Private Function ChkPfx(A() As Lnx, OEr$()) As Lnx()
Dim J%
For J = 0 To UB(A)
    If FstChr(A(J).Lin) <> "?" Then
        PushI OEr, MsgPfx(A(J))
    Else
        PushObj ChkPfx, A(J)
    End If
Next
End Function

Private Sub Evl(A() As SwBrk, OStmt As Dictionary, OFld As Dictionary)
Dim A1() As SwBrk
Dim A2() As SwBrk
Dim Sw As New Dictionary
Dim OSw As New Dictionary
Dim IsEvl As Boolean
Dim I%, J%
IsEvl = True
A1 = A
While IsEvl
    IsEvl = False
    J = J + 1
    If J > 1000 Then Stop
    For I = 0 To UB(A1)
        If EvlLin(A1(I), Sw, OSw) Then
            IsEvl = True
        Else
            PushObj A2, A1(I)
        End If
    Next
    If False Then
        Brw AyAddAp( _
            LblTabAyFmt("*A1", SwBrkAyFmt(A1)), _
            LblTabAyFmt("*A2", SwBrkAyFmt(A2)), _
            LblTabAyFmt("OSw", DicFmt(OSw)))
        Stop
    End If
    A1 = A2
    Erase A2
    Set Sw = DicClone(OSw)
Wend
If Sz(OSw) > 0 Then Stop
EvlSwAsg OSw, OStmt, OFld
End Sub

Private Function EvlBoolTerm(A, Sw As Dictionary, ORslt As Boolean) As Boolean
If A_Pm.Exists(A) Then
    ORslt = A_Pm(A)
    EvlBoolTerm = True
    Exit Function
End If
If Not Sw.Exists(A) Then Exit Function
ORslt = Sw(A)
EvlBoolTerm = True
End Function

Private Function EvlLin(A As SwBrk, Sw As Dictionary, OSw As Dictionary) As Boolean
'Return True and set Result if evalulated
Const CSub$ = CMod & "Evl"
If Sw.Exists(A.Nm) Then Er CSub, "[SwBrk] should not be found in [Sw]", SwBrkStr(A), DicFmt(Sw)
Dim Ay$(): Ay = A.TermAy
Dim ORslt As Boolean, IsEvl As Boolean
Select Case A.OpStr
Case "OR":  IsEvl = EvlTermAy(Ay, "OR", Sw, ORslt)
Case "AND": IsEvl = EvlTermAy(Ay, "AND", Sw, ORslt)
Case "NE":  IsEvl = EvlT1T2(Ay(0), Ay(1), "NE", Sw, ORslt)
Case "EQ":  IsEvl = EvlT1T2(Ay(0), Ay(1), "EQ", Sw, ORslt)
Case Else: Er CSub, "[SwBrk] has invalid [OpStr], where [Valid OpStr]", SwBrkStr(A), A.OpStr, "OR AND NE EQ"
End Select
If IsEvl Then
    OSw.Add A.Nm, ORslt
    EvlLin = True
End If
End Function

Private Sub EvlSwAsg(A As Dictionary, OStmt As Dictionary, OFld As Dictionary)
Set OStmt = New Dictionary
Set OFld = New Dictionary
Dim K
For Each K In A.Keys
    If FstTwoChr(K) <> "?#" Then  ' Skip.  It is temp-Sw
        Select Case Left(K, 5)
        Case "?SEL#", "?UPD#"     ' It is StmtSw
            OStmt.Add Mid(K, 5), A(K)
        Case Else                 ' It is FldSw
            OFld.Add K, A(K)
        End Select
    End If
Next
End Sub

Private Function EvlT1(A, Sw As Dictionary, ORslt$) As Boolean
EvlT1 = EvlTerm(A, Sw, ORslt)
End Function

Private Function EvlT1T2(T1$, T2$, EQ_NE$, Sw As Dictionary, ORslt As Boolean) As Boolean
'Return True and set ORslt if evaluated
Const CSub$ = CMod & "EvlT1T2"
Dim S1$, S2$
If Not EvlT1(T1, Sw, S1) Then Exit Function
If Not EvlT2(T2, Sw, S2) Then Exit Function
Select Case EQ_NE
Case "EQ": ORslt = S1 = S2: EvlT1T2 = True
Case "NE": ORslt = S1 <> S2:: EvlT1T2 = True
Case Else: Er CSub, "[EQ_NE] does not eq EQ or NE", EQ_NE
End Select
End Function

Private Function EvlT2(A, Sw As Dictionary, ORslt$) As Boolean
'Return True is evalulated
'switch-term begins with @ or ? or it is *Blank.  @ is for parameter & ? is for switch
'  If @, it will evaluated to str by lookup from Pm
'        if not exist in {Pm}, stop, it means the validation fail to remove this term
'  If ?, it will evaluated to bool by lookup from Sw
'        if not exist in {Sw}, return None
'  Otherwise, just return SomVar(A)
If EvlTerm(A, Sw, ORslt) Then Exit Function
If FstChr(A) = "*" Then
    If UCase(A) <> "*BLANK" Then Stop ' it means the validation fail to remove this term
    ORslt = ""
Else
    ORslt = A
End If
EvlT2 = True
End Function

Private Function EvlTerm(A, Sw As Dictionary, ORslt$) As Boolean
If A_Pm.Exists(A) Then
    ORslt = A_Pm(A)
    EvlTerm = True
    Exit Function
End If
If Not Sw.Exists(A) Then Exit Function
ORslt = Sw(A)
EvlTerm = True
End Function

Private Function EvlTermAy1(A$(), AND_OR, Sw As Dictionary) As Boolean()
Dim Rslt As Boolean, O() As Boolean, I
For Each I In A
    If Not EvlBoolTerm(I, Sw, Rslt) Then Exit Function
    PushI O, Rslt
Next
EvlTermAy1 = O
End Function

Private Function EvlTermAy(A$(), AND_OR$, Sw As Dictionary, ORslt As Boolean) As Boolean
If Sz(A) = 0 Then Stop
Dim BoolAy() As Boolean
    BoolAy = EvlTermAy1(A, AND_OR, Sw)
    If Sz(BoolAy) = 0 Then Exit Function
    
Select Case AND_OR
Case "AND": ORslt = BoolAyIsAllTrue(BoolAy)
Case "OR":  ORslt = BoolAyIsSomTrue(BoolAy)
Case Else: Stop
End Select
EvlTermAy = True
End Function

Private Function FndEr(A() As Lnx) As Variant()
Dim B() As SwBrk
    B = SwBrkAy(A)
'    Er = AyAddAp(ChkLin(B), ChkDupNm(B), ChkFld(B), ChkLeftOvr(B))
'FndEr = Array(Er, B)
End Function

Sub FndSwAsg(A() As Lnx, Pm As Dictionary, OStmtSw As Dictionary, OFldSw As Dictionary, OEr$())
Set A_Pm = Pm
Dim B() As SwBrk
AyAsg FndEr(A), OEr, B
Evl B, OStmtSw, OFldSw
End Sub

Function ChkLeftOvr(IO() As SwBrk) As String()
End Function

Private Function Msg$(A As SwBrk, B$)
Msg = SwBrkStr(A) & " --- " & B
End Function

Private Function MsgDupNm(A() As SwBrk) As String()
Dim I
For Each I In AyNz(A)
    PushI MsgDupNm, Msg(CvSwBrk(I), "Dup name")
Next
End Function

Private Function MsgLeftOvrAftEvl(A() As SwBrk, Sw As Dictionary) As String()
If Sz(A) = 0 Then Exit Function
Dim I
PushI MsgLeftOvrAftEvl, "Following lines cannot be further evaluated:"
For Each I In A
    PushI MsgLeftOvrAftEvl, vbTab & SwBrkStr(CvSwBrk(I))
Next
PushIAy MsgLeftOvrAftEvl, DicLblLy(Sw, "Following is the [Sw] after evaluated:")
End Function

Private Function MsgNoNm$(A As SwBrk)
MsgNoNm = Msg(A, "No name")
End Function

Private Function MsgOpStrEr$(A As SwBrk)
MsgOpStrEr = Msg(A, "2nd Term [Op] is invalid operator.  Valid operation [NE EQ AND OR]")
Stop
End Function

Private Function MsgPfx$(A As SwBrk)
MsgPfx = Msg(A, "First Char must be @")
End Function

Private Function MsgTermCntAndOr$(A As SwBrk)
MsgTermCntAndOr = Msg(A, "When 2nd-Term (Operator) is [AND OR], at least 1 term")
End Function

Private Function MsgTermCntEqNe$(A As SwBrk)
MsgTermCntEqNe = Msg(A, "When 2nd-Term (Operator) is [EQ NE], only 2 terms are allowed")
End Function

Private Function MsgTermMustBegWithQuestOrAt$(TermAy$(), A As SwBrk)
MsgTermMustBegWithQuestOrAt = Msg(A, "Terms[" & JnSpc(TermAy) & "] must begin with either [?] or [@?]")
End Function

Private Function MsgTermNotInPm$(TermAy$(), A As SwBrk)
MsgTermNotInPm = Msg(A, "Terms[" & JnSpc(TermAy) & "] begin with [@?] must be found in Pm")
End Function

Private Function MsgTermNotInSw$(TermAy$(), A As SwBrk, SwNm As Dictionary)
MsgTermNotInSw = Msg(A, "Terms[" & JnSpc(TermAy) & "] begin with [?] must be found in Switch")
End Function

Private Function OpStrAy() As String()
Static X$()
If Sz(X) = 0 Then X = SslSy("OR AND NE EQ")
OpStrAy = X
End Function

Private Function PmNy() As String()
PmNy = DicStrKy(A_Pm)
End Function

Private Function SamplePm() As Dictionary
Set SamplePm = LyDic(SamplePmLy)
End Function

Private Function SamplePmLy() As String()
PushI SamplePmLy, "@?BrkMbr 0"
PushI SamplePmLy, "@?BrkSto 0"
PushI SamplePmLy, "@?BrkCrd 0"
PushI SamplePmLy, "@?BrkDiv 0"
'-- @XXX means txt and optional, allow, blank
PushI SamplePmLy, "@SumLvl  Y"
PushI SamplePmLy, "@?MbrEmail 1"
PushI SamplePmLy, "@?MbrNm    1"
PushI SamplePmLy, "@?MbrPhone 1"
PushI SamplePmLy, "@?MbrAdr   1"
'-- @@ mean compulasary
PushI SamplePmLy, "@DteFm 20170101"
'@@DteTo 20170131
PushI SamplePmLy, "@LisDiv 1 2"
PushI SamplePmLy, "@LisSto"
PushI SamplePmLy, "@LisCrd"
PushI SamplePmLy, "@CrdExpr ..."
PushI SamplePmLy, "@CrdExpr ..."
PushI SamplePmLy, "@CrdExpr ..."
End Function

Private Function SampleSwLnxAy() As Lnx()
Dim M$()
PushI M, "?#LvlY    EQ @SumLvl Y"
PushI M, "?#LvlY    EQ @SumLvl Y"
PushI M, "?#LvlM    EQ @SumLvl M"
PushI M, "?#LvlW    EQ @SumLvl W"
PushI M, "?#LvlD    EQ @SumLvl D"
PushI M, "?Y       OR ?#LvlD ?#LvlW ?#LvlM ?#LvlY"
PushI M, "?M       OR ?#LvlD ?#LvlW ?#LvlM"
PushI M, "?W       OR ?#LvlD ?#LvlW"
PushI M, "?D       OR ?#LvlD"
PushI M, "?Dte     OR ?#LvlD"
PushI M, "?Mbr     OR @?BrkMbr XX"
PushI M, "?MbrCnt  OR @?BrkMbr"
PushI M, "?Div     OR @?BrkDiv"
PushI M, "?Sto     OR @?BrkSto"
PushI M, "?Crd     OR @?BrkCrd"
PushI M, "?SEL#Div NE @LisDiv *blank"
PushI M, "?SEL#Sto NE @LisSto *blank"
PushI M, "?SEL#Crd NE @LisCrd *blank"
SampleSwLnxAy = LyLnxAy(M)
End Function

Private Function SwBrk(A As Lnx) As SwBrk
Const CSub$ = CMod & "SwBrk"
Dim L$, Ix%, OEr$()
L = A.Lin
Ix = A.Ix
If LinIsRmkLin(L) Then Er CSub, "[SwLin], [Ix] is a remark line.  It should be removed before calling Evl", A.Lin, A.Ix
Set SwBrk = New SwBrk
With SwBrk
    .Nm = ShfTerm(L)
    .OpStr = UCase(ShfTerm(L))
    .TermAy = SslSy(L)
    .Ix = Ix
End With
End Function

Private Function SwBrkAy(A() As Lnx) As SwBrk()
Dim I
For Each I In AyNz(A)
    PushObj SwBrkAy, SwBrk(CvLnx(I))
Next
End Function

Private Sub SwBrkAyBrw(A() As SwBrk)
AyBrw SwBrkAyFmt(A)
End Sub

Private Function SwBrkAyFmt(A() As SwBrk) As String()
Dim I
For Each I In AyNz(A)
    PushI SwBrkAyFmt, SwBrkStr(CvSwBrk(I))
Next
End Function

Private Function SwBrkStr$(A As SwBrk)
With A
    SwBrkStr = Quote(.Ix, "L#(*) ") & QuoteSqBkt(JnSpc(Array(.Nm, .OpStr, JnSpc(.TermAy))))
End With
End Function

Private Function SwNmDic(A() As SwBrk) As Dictionary
Set SwNmDic = AyIxDic(SwNy(A))
End Function

Private Function SwNy(A() As SwBrk) As String()
Dim J%
For J = 0 To UB(A)
    PushI SwNy, A(J).Nm
Next
End Function
Sub AAA()
Z_FndSwAsg
End Sub
Private Sub Z_FndSwAsg()
Dim Stmt As Dictionary, Fld As Dictionary, Er$()
FndSwAsg SampleSwLnxAy, SamplePm, Stmt, Fld, Er
Brw NmssApLy("SwLnxAy Pm StmtSw FldSw", _
    LnxAyFmt(SampleSwLnxAy), _
    A_Pm, Stmt, Fld)
End Sub

