Attribute VB_Name = "MVb_Ay"
Option Explicit
Sub AyAsg(A, ParamArray OAp())
Dim Av(): Av = OAp
Dim J%
For J = 0 To Min(UB(Av), UB(A))
    Asg A(J), OAp(J)
Next
End Sub

Function AyAsgAy(A, OIntoAy)
If TypeName(A) = TypeName(OIntoAy) Then
    OIntoAy = A
    AyAsgAy = OIntoAy
    Exit Function
End If
If Sz(A) = 0 Then
    Erase OIntoAy
    AyAsgAy = OIntoAy
    Exit Function
End If
Dim U&
    U = UB(A)
ReDim OIntoAy(U)
Dim I, J&
For Each I In A
    Asg I, OIntoAy(J)
    J = J + 1
Next
AyAsgAy = OIntoAy
End Function

Sub AyAsgT1AyRestAy(A, OT1Ay$(), ORestAy$())
Dim U&, J&
U = UB(A)
If U = -1 Then
    Erase OT1Ay, ORestAy
    Exit Sub
End If
ReDim OT1Ay(U)
ReDim ORestAy(U)
For J = 0 To U
    BrkAsg A(J), " ", OT1Ay(J), ORestAy(J)
Next
End Sub

Sub AyAssSamSz(A, B)
Ass Sz(A) = Sz(B)
End Sub

Function AyBrw(A, Optional Fnn$)
Dim T$
T = TmpFt("AyBrw", Fnn)
AyWrt A, T
FtBrw T
End Function

Sub AyChkEq(Ay1, Ay2, Optional Nm1$ = "Ay1", Optional Nm2$ = "Ay2")
Chk AyEqChk(Ay1, Ay2, Nm1, Nm2)
End Sub

Function AyCln(A)
If Not IsArray(A) Then Er "AyCln", "Given [A] is not an array, but [TypeName]", A, TypeName(A)
Dim O
O = A
Erase O
AyCln = O
End Function

Function AyCsv$(A)
If Sz(A) = 0 Then Exit Function
Dim O$(), U&, J&, V
U = UB(A)
ReDim O(U)
For Each V In A
    O(J) = VarCsv(V)
    J = J + 1
Next
AyCsv = Join(O, ",")
End Function
Function AyDupChk(A, QMsg$) As String()
Dim Dup
Dup = AyWhDup(A)
If Sz(Dup) = 0 Then Exit Function
Push AyDupChk, FmtQQ(QMsg, JnSpc(Dup))
End Function

Function AyDupT1(A) As String()
AyDupT1 = AyWhDup(AyTakT1(A))
End Function

Function AyEmpChk(A, Msg$) As String()
If Sz(A) = 0 Then AyEmpChk = Sy(Msg)
End Function

Function AyEqChk(Ay1, Ay2, Optional Ay1Nm$ = "Exp", Optional Ay2Nm$ = "Act") As String()
Dim U&: U = UB(Ay1)
Dim O$()
    If U <> UB(Ay2) Then Push O, FmtQQ("Array [?] and [?] has different Sz: [?] [?]", Ay1Nm, Ay2Nm, Sz(Ay1), Sz(Ay2)): GoTo X
If Sz(Ay1) = 0 Then Exit Function
Dim O1$()
    Dim A2: A2 = Ay2
    Dim J&, ReachLimit As Boolean
    Dim Cnt%
    For J = 0 To U
        If Ay1(J) <> Ay2(J) Then
            Push O1, FmtQQ("[?]-th Ele is diff: ?[?]<>?[?]", Ay1Nm, Ay2Nm, Ay1(J), Ay2(J))
            Cnt = Cnt + 1
        End If
        If Cnt > 10 Then
            ReachLimit = True
            Exit For
        End If
    Next
If IsEmp(O1) Then Exit Function
Dim O2$()
    Push O2, FmtQQ("Array [?] and [?] both having size[?] have differnt element(s):", Ay1Nm, Ay2Nm, Sz(Ay1))
    If ReachLimit Then
        Push O2, FmtQQ("At least [?] differences:", Sz(O1))
    End If
PushAy O, O2
PushAy O, O1
X:
Push O, FmtQQ("Ay-[?]:", Ay1Nm)
PushAy O, AyQuote(Ay1, "[]")
Push O, FmtQQ("Ay-[?]:", Ay2Nm)
PushAy O, AyQuote(Ay2, "[]")
AyEqChk = O
End Function

Function AyOfAy_Ay(AyOfAy)
If Sz(AyOfAy) = 0 Then Exit Function
Dim O
O = AyCln(AyOfAy(0))
Dim X
For Each X In AyOfAy
    PushAy O, X
Next
AyOfAy_Ay = O
End Function
Private Sub Z_AyFlat()
Dim AyOfAy()
AyOfAy = Array(SslSy("a b c d"), SslSy("a b c"))
Ept = SslSy("a b c d a b c")
GoSub Tst
Exit Sub
Tst:
    Act = AyFlat(AyOfAy)
    C
    Return
End Sub

Function AyFlat(AyOfAy)
AyFlat = AyOfAy_Ay(AyOfAy)
End Function

Function AyItmCnt%(A, M)
If Sz(A) = 0 Then Exit Function
Dim O%, X
For Each X In AyNz(A)
    If X = M Then O = O + 1
Next
AyItmCnt = O
End Function

Function AyKeepLasN(A, N)
Dim O, J&, I&, U&, Fm&, NewU&
U = UB(A)
If U < N Then AyKeepLasN = A: Exit Function
O = A
Fm = U - N + 1
NewU = N - 1
For J = Fm To U
    Asg O(J), O(I)
    I = I + 1
Next
ReDim Preserve O(NewU)
AyKeepLasN = O
End Function

Function AyLasEle(A)
Asg A(UB(A)), AyLasEle
End Function

Function AyLines$(A)
AyLines = JnCrLf(A)
End Function

Function AyMid(A, Fm, Optional L = 0)
AyMid = AyCln(A)
Dim J&
Dim E&
    Select Case True
    Case L = 0: E = UB(A)
    Case Else:  E = Min(UB(A), L + Fm - 1)
    End Select
For J = Fm To E
    Push AyMid, A(J)
Next
End Function


Function AyNPfxStar%(A)
Dim O%, X
For Each X In AyNz(A)
    If FstChr(X) = "*" Then AyNPfxStar = O: Exit Function
    O = O + 1
Next
End Function
Function AyNxtNm$(A, Nm$, Optional MaxN% = 99)
If Not AyHas(A, Nm) Then AyNxtNm = Nm: Exit Function
Dim J%, O$
For J = 1 To MaxN
    O = Nm & Format(J, "00")
    If Not AyHas(A, O) Then AyNxtNm = O: Exit Function
Next
Stop
End Function

Function AyNz(A)
If Sz(A) = 0 Then Set AyNz = New Collection Else AyNz = A
End Function

Sub AyPushMsgAv(A, Msg$, Av())
PushAy A, MsgAvLy(Msg, Av)
End Sub


Function AyRTrim(A$()) As String()
If Sz(A) = 0 Then Exit Function
Dim O$(), I
For Each I In A
    Push O, RTrim(I)
Next
AyRTrim = O
End Function

Function AyReSz(A, SzAy) ' return Ay as [A] with sz as [SzAy]
If Sz(SzAy) = 0 Then AyReSz = A: Exit Function
Dim O
O = A
ReDim O(UB(SzAy))
AyReSz = O
End Function

Function AyReverseI(A)
Dim O: O = A
Dim J&, U&
U = UB(O)
For J = 0 To U
    O(J) = A(U - J)
Next
AyReverseI = O
End Function

Function AyReverseObj(A)
Dim O: O = A
Dim J&, U&
U = UB(O)
For J = 0 To U
    Set O(J) = A(U - J)
Next
AyReverseObj = O
End Function

Function AyRplFTIx(A, B As FTIx, AySeg)
AyRplFTIx = AyRplSeg(A, B.FmIx, B.ToIx, AySeg)
End Function

Function AyRplSeg(A, FmIx&, ToIx&, AySeg)
Dim B()
    B = AyBrkInto3Ay(A, FmIx, ToIx)
AyRplSeg = B(1)
    PushAy AyRplSeg, AySeg
    PushAy AyRplSeg, B(2)
End Function

Function AyRplStar(A$(), By) As String()
Dim X
For Each X In AyNz(A)
    PushI AyRplStar, Replace(X, By, "*")
Next
End Function

Function AyRplT1(A$(), T1$) As String()
AyRplT1 = AyAddPfx(AyRmvT1(A), T1 & " ")
End Function

Function AySampleLin$(A)
Dim S$, U&
U = UB(A)
If U >= 0 Then
    Select Case True
    Case IsPrim(A(0)): S = "[" & A(0) & "]"
    Case IsObject(A(0)), IsArray(A(0)): S = "[*Ty:" & TypeName(A(0)) & "]"
    Case Else: Stop
    End Select
End If
AySampleLin = "*Ay:[" & U & "]" & S
End Function

Function AySel(A, M) As Boolean
If Sz(A) = 0 Then AySel = True: Exit Function
AySel = AyHas(A, M)
End Function


Function AySeqCntDic(A) As Dictionary 'The return dic of key=AyEle pointing to Long(1) with Itm0 as Seq# and Itm1 as Cnt
Dim S&, O As New Dictionary, L&(), X
For Each X In AyNz(A)
    If O.Exists(X) Then
        L = O(X)
        L(1) = L(1) + 1
        O(X) = L
    Else
        ReDim L(1)
        L(0) = S
        L(1) = 1
        O.Add X, L
    End If
Next
Set AySeqCntDic = O
End Function
Function AySqH(A) As Variant()
Dim N&: N = Sz(A)
If N = 0 Then Exit Function
Dim J&, V
Dim O()
ReDim O(1 To 1, 1 To N)
For Each V In A
    J = J + 1
    O(1, J) = V
Next
AySqH = O
End Function

Function AySqV(A) As Variant()
Dim N&: N = Sz(A)
If N = 0 Then Exit Function
Dim J&, V
Dim O()
ReDim O(1 To N, 1 To 1)
For Each V In A
    J = J + 1
    O(J, 1) = V
Next
AySqV = O
End Function

Function AyT1Chd(A, T1) As String()
AyT1Chd = AyRmvT1(AyWhT1(A, T1))
End Function

Function AyTrim(A) As String()
Dim X
For Each X In AyNz(A)
    Push AyTrim, Trim(X)
Next
End Function

Function AyWdt%(A)
Dim O%, J&
For J = 0 To UB(A)
    O = Max(O, Len(A(J)))
Next
AyWdt = O
End Function

Function AyWrpPad(A, W%) As String() ' Each Itm of Ay[A] is padded to line with AyWdt(A).  return all padded lines as String()
Dim O$(), X, I%
ReDim O(0)
For Each X In AyNz(A)
    If Len(O(I)) + Len(X) < W Then
        O(I) = O(I) & X
    Else
        PushI O, X
        I = I + 1
    End If
Next
AyWrpPad = O
End Function

Sub AyWrt(A, Ft)
StrWrt JnCrLf(A), Ft
End Sub

Function AyZip(A1, A2) As Variant()
Dim U1&: U1 = UB(A1)
Dim U2&: U2 = UB(A2)
Dim U&: U = Max(U1, U2)
Dim O(), J&
ReSz O, U
For J = 0 To U
    If U1 >= J Then
        If U2 >= J Then
            O(J) = Array(A1(J), A2(J))
        Else
            O(J) = Array(A1(J), Empty)
        End If
    Else
        If U2 >= J Then
            O(J) = Array(, A2(J))
        Else
            Stop
        End If
    End If
Next
AyZip = O
End Function

Function AyZipAp(A1, ParamArray Ap()) As Variant()
Dim Av(): Av = Ap
Dim UCol%
    UCol = UB(Av)

Dim URow1&
    URow1 = UB(A1)

Dim URow&
Dim URowAy&()
    Dim J%, IURow%
    URow = URow1
    For J = 0 To UB(Av)
        IURow = UB(Av(J))
        Push URowAy, IURow
        If IURow > URow Then URow = IURow
    Next

Dim ODry()
    Dim Dr()
    ReSz ODry, URow
    Dim I%
    For J = 0 To URow
        Erase Dr
        If URow1 >= J Then
            Push Dr, A1(J)
        Else
            Push Dr, Empty
        End If
        For I = 0 To UB(Av)
            If URowAy(I) >= J Then
                Push Dr, Av(I)(J)
            Else
                Push Dr, Empty
            End If
        Next
        ODry(J) = Dr
    Next
AyZipAp = ODry
End Function

Function Aydic_to_KeyCntMulItmColDry(A As Dictionary) As Variant()
If A.Count = 0 Then Exit Function
Dim O(), K, Dr(), Ay, J&
ReDim O(A.Count - 1)
For Each K In A.Keys
    Ay = A(K): If Not IsArray(Ay) Then Stop
    O(J) = AyIns2(Ay, K, Sz(Ay))
    J = J + 1
Next
Aydic_to_KeyCntMulItmColDry = O
End Function

Function ItmAddAy(Itm, Ay)
ItmAddAy = AyIns(Ay, Itm)
End Function

Function SSAySy(SSAy) As String()
Dim SS
For Each SS In AyNz(SSAy)
    PushIAy SSAySy, SslSy(SS)
Next
End Function

Function VyDr(A(), FF, Fny$()) As Variant()
Dim IxAy&(), U%
    IxAy = AyIxAy(Fny, SslSy(FF))
    U = AyMax(IxAy)
    GoSub X_ChkIxAy
Dim O(), J%, Ix
ReDim O(U)
For Each Ix In IxAy
    O(Ix) = A(J)
    J = J + 1
Next
VyDr = O
Exit Function
X_ChkIxAy:
    For Each Ix In IxAy
        If Ix <= -1 Then Stop
    Next
    Return
End Function

Private Sub ZZZ_AyBrkInto3Ay()
Dim A(): A = Array(1, 2, 3, 4)
Dim Act(): Act = AyBrkInto3Ay(A, 1, 2)
Ass Sz(Act) = 3
Ass IsEqAy(Act(0), Array(1))
Ass IsEqAy(Act(1), Array(2, 3))
Ass IsEqAy(Act(2), Array(4))
End Sub

Sub ZZ_AyAsgAp()
Dim O%, A$
'AyAsgAp Array(234, "abc"), O, A
Ass O = 234
Ass A = "abc"
End Sub

Private Sub ZZ_AyEqChk()
AyDmp AyEqChk(Array(1, 2, 3, 3, 4), Array(1, 2, 3, 4, 4))
End Sub



Private Sub ZZ_AyMax()
Dim A()
Dim Act
Act = AyMax(A)
Stop
End Sub

Private Sub ZZ_AyMinus()
Dim Act(), Exp()
Dim Ay1(), Ay2()
Ay1 = Array(1, 2, 2, 2, 4, 5)
Ay2 = Array(2, 2)
Act = AyMinus(Ay1, Ay2)
Exp = Array(1, 2, 4, 5)
AyChkEq Exp, Act
'
Act = AyMinusAp(Array(1, 2, 2, 2, 4, 5), Array(2, 2), Array(5))
Exp = Array(1, 2, 4)
AyChkEq Exp, Act
End Sub

Private Sub ZZ_AyRmvEmpEleAtEnd()
Dim A: A = Array(Empty, Empty, Empty, 1, Empty, Empty)
Dim Act: Act = AyRmvEmpEleAtEnd(A)
Ass Sz(Act) = 4
Ass Act(3) = 1
End Sub

Private Sub ZZ_AySy()
Dim Act$(): Act = AySy(Array(1, 2, 3))
Ass Sz(Act) = 3
Ass Act(0) = 1
Ass Act(1) = 2
Ass Act(2) = 3
End Sub

Private Sub ZZ_AyTrim()
AyDmp AyTrim(Array(1, 2, 3, "  a"))
End Sub

Sub Z()
Z_AyDupChk
Z_AyIns
Z_AyInsAy
Z_Aydic_to_KeyCntMulItmCol_Dry
Z_VyDr
End Sub

Private Sub Z_AyDupChk()
Dim Ay
Ay = Array("1", "1", "2")
Ept = ApSy("This item[1] is duplicated")
GoSub Tst
Exit Sub
Tst:
    Act = AyDupChk(Ay, "This item[?] is duplicated")
    C
    Return
End Sub

Private Sub Z_AyEqChk()
AyDmp AyEqChk(Array(1, 2, 3, 3, 4), Array(1, 2, 3, 4, 4))
End Sub

Sub Z_AyFTIxBrk()
Dim A(): A = Array(1, 2, 3, 4)
Dim M As FTIx: M = FTIx(1, 2)
Dim Act(): Act = AyFTIxBrk(A, M)
Ass Sz(Act) = 2
Ass IsEqAy(Act(0), Array(1))
Ass IsEqAy(Act(1), Array(2, 3))
Ass IsEqAy(Act(2), Array(4))
End Sub

Sub Z_AyHasDupEle()
Ass AyHasDupEle(Array(1, 2, 3, 4)) = False
Ass AyHasDupEle(Array(1, 2, 3, 4, 4)) = True
End Sub

Sub AyIxAyAsg(Dr, IxAy%(), ParamArray OAp())
Dim J%
For J = 0 To UB(IxAy)
    If IsObject(OAp(J)) Then
        Set OAp(J) = Dr(IxAy(J))
    Else
        OAp(J) = Dr(IxAy(J))
    End If
Next
End Sub
Private Sub Z_AyIns()
Dim A, M, At&
'
A = Array(1, 2, 3)
M = "X"
Ept = Array("X", 1, 2, 3)
GoSub Tst
'
Exit Sub
Tst:
    Act = AyIns(A, M, At)
    C
Return
End Sub

Private Sub Z_AyInsAy()
Dim Act, Exp, A(), B(), At&
A = Array(1, 2, 3, 4)
B = Array("X", "Z")
At = 1
Exp = Array(1, "X", "Z", 2, 3, 4)

Act = AyInsAy(A, B, At)
Ass IsEqAy(Act, Exp)
End Sub

Sub Z_AyMinus()
Dim Act(), Exp()
Dim Ay1(), Ay2()
Ay1 = Array(1, 2, 2, 2, 4, 5)
Ay2 = Array(2, 2)
Act = AyMinus(Ay1, Ay2)
Exp = Array(1, 2, 4, 5)
AyabEqChk Exp, Act
'
Act = AyMinusAp(Array(1, 2, 2, 2, 4, 5), Array(2, 2), Array(5))
Exp = Array(1, 2, 4)
AyabEqChk Exp, Act
End Sub

Sub Z_AySy()
Dim Act$(): Act = AySy(Array(1, 2, 3))
Ass Sz(Act) = 3
Ass Act(0) = 1
Ass Act(1) = 2
Ass Act(2) = 3
End Sub

Private Sub Z_AyTrim()
AyDmp AyTrim(ApSy(1, 2, 3, "  a"))
End Sub

Private Sub Z_Aydic_to_KeyCntMulItmCol_Dry()
Dim A As New Dictionary, Act()
A.Add "A", Array(1, 2, 3)
A.Add "B", Array(2, 3, 4)
A.Add "C", Array()
A.Add "D", Array("X")
Act = Aydic_to_KeyCntMulItmColDry(A)
Ass Sz(Act) = 4
Ass IsEqAy(Act(0), Array("A", 3, 1, 2, 3))
Ass IsEqAy(Act(1), Array("B", 3, 2, 3, 4))
Ass IsEqAy(Act(2), Array("C", 0))
Ass IsEqAy(Act(3), Array("D", 1, "X"))
End Sub
Private Sub Z_VyDr()
Dim Fny$(), FF, Vy()
Fny = SslSy("A B C D E f")
FF = "C E"
Vy = Array(1, 2)
Ept = Array(Empty, Empty, 1, Empty, 2)
GoSub Tst
Exit Sub
Tst:
    Act = VyDr(Vy, FF, Fny)
    C
    Return
End Sub
