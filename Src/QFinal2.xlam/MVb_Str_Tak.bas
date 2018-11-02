Attribute VB_Name = "MVb_Str_Tak"
Option Explicit
Sub Z()
Z_TakAftBkt
Z_TakBet
Z_TakBetBkt
End Sub

Function TakBefDot$(A)
TakBefDot = TakBef(A, ".")
End Function

Function TakAft$(S, Sep, Optional NoTrim As Boolean)
TakAft = Brk1(S, Sep, NoTrim).S2
End Function

Function TakAftAt$(A, At&, S)
If At = 0 Then Exit Function
TakAftAt = Mid(A, At + Len(S))
End Function


Function TakAftDotOrAll$(A)
TakAftDotOrAll = TakAftOrAll(A, ".")
End Function

Function TakAftMust$(A, Sep, Optional NoTrim As Boolean)
TakAftMust = Brk(A, Sep, NoTrim).S2
End Function

Function TakAftOrAll$(S, Sep, Optional NoTrim As Boolean)
TakAftOrAll = Brk2(S, Sep, NoTrim).S2
End Function

Function TakAftOrAllRev$(A, S)
TakAftOrAllRev = StrDft(TakAftRev(A, S), A)
End Function

Function TakAftRev$(S, Sep, Optional NoTrim As Boolean)
TakAftRev = Brk1Rev(S, Sep, NoTrim).S2
End Function

Function TakBef$(S, Sep, Optional NoTrim As Boolean)
TakBef = Brk2(S, Sep, NoTrim).S1
End Function

Function TakBefAt(A, At&)
If At = 0 Then Exit Function
TakBefAt = Left(A, At - 1)
End Function

Function TakBefDD$(A)
TakBefDD = RTrim(TakBefOrAll(A, "--"))
End Function

Function TakBefDDD$(A)
TakBefDDD = RTrim(TakBefOrAll(A, "---"))
End Function

Function TakBefMust$(S, Sep$, Optional NoTrim As Boolean)
TakBefMust = Brk(S, Sep, NoTrim).S1
End Function

Function TakBefOrAll$(S, Sep, Optional NoTrim As Boolean)
TakBefOrAll = Brk1(S, Sep, NoTrim).S1
End Function

Function TakBefOrAllRev$(A, S)
TakBefOrAllRev = StrDft(TakBefRev(A, S), A)
End Function

Function TakBefRev$(S, Sep, Optional NoTrim As Boolean)
TakBefRev = BrkRev(S, Sep, NoTrim).S1
End Function

Function TakBet$(S, S1, S2, Optional NoTrim As Boolean, Optional InclMarker As Boolean)
With Brk1(S, S1, NoTrim)
   If .S2 = "" Then Exit Function
   Dim O$: O = Brk1(.S2, S2, NoTrim).S1
   If InclMarker Then O = S1 & O & S2
   TakBet = O
End With
End Function

Private Sub Z_TakBetBkt()
Dim Act$
   Dim S$
   S = "sdklfjdsf(1234()567)aaa("
   Act = TakBetBkt(S)
   Ass Act = "1234()567"
End Sub

Function TakNm$(A)
Dim J%
If Not IsLetter(Left(A, 1)) Then Exit Function
For J = 2 To Len(A)
    If Not IsNmChr(Mid(A, J, 1)) Then
        TakNm = Left(A, J - 1)
        Exit Function
    End If
Next
TakNm = A
End Function

Function TakPfx$(Lin, Pfx$) ' Return [Pfx] if [Lin] has such pfx else return ""
If HasPfx(Lin, Pfx) Then TakPfx = Pfx
End Function

Function PfxAyFst$(PfxAy$(), Lin) ' Return Fst ele-P of [PfxAy] if [Lin] has pfx ele-P
Dim P
For Each P In PfxAy
    If HasPfx(Lin, P) Then PfxAyFst = P: Exit Function
Next
End Function

Function PfxAyFstSpc$(PfxAy$(), Lin) ' Return Fst ele-P of [PfxAy] if [Lin] has pfx ele-P and a space
Dim P
For Each P In PfxAy
    If HasPfx(Lin, P & " ") Then PfxAyFstSpc = P: Exit Function
Next
End Function

Function TakPfxAy$(Lin, PfxAy$()) ' Return Fst ele-P of [PfxAy] if [Lin] has pfx ele-P
TakPfxAy = PfxAyFst$(PfxAy, Lin)
End Function

Function TakPfxAySpc$(Lin, PfxAy$()) ' Return Fst ele-P of [PfxAy] if [Lin] has pfx ele-P and a space
TakPfxAySpc = PfxAyFstSpc$(PfxAy, Lin)
End Function

Function TakPfxS$(Lin, Pfx$) ' Return [Pfx] if [Lin] has such pfx+" " else return ""
If HasPfx(Lin, Pfx) Then If Mid(Lin, Len(Pfx) + 1, 1) = " " Then TakPfxS = Pfx
End Function

Function TakT1$(A)
If FstChr(A) <> "[" Then TakT1 = TakBef(A, " "): Exit Function
Dim P%
P = InStr(A, "]")
If P = 0 Then Stop
TakT1 = Mid(A, 2, P - 2)
End Function

Private Sub Z_TakAftBkt()
Dim A$
A = "(lsk(aa)df lsdkfj) A"
Ept = " A"
GoSub Tst
Exit Sub
Tst:
    Act = TakAftBkt(A)
    C
    Return
End Sub

Private Sub Z_TakBet()
Dim Lin$
Lin = "Excel 8.0;HDR=YES;IMEX=2;DATABASE=??       | DATABASE= | ; | ??":            GoSub Tst
Lin = "Excel 8.0;HDR=YES;IMEX=2;DATABASE=??;AA=XX | DATABASE= | ; | ??":            GoSub Tst
Lin = "lkjsdf;dkfjl;Data Source=Johnson;lsdfjldf  | Data Source= | ; | Johnson":    GoSub Tst
Exit Sub
Tst:
    Dim FmStr$, ToStr$
    AyAsg AyTrim(SplitVBar(Lin)), Lin, FmStr, ToStr, Ept
    Act = TakBet(Lin, FmStr, ToStr)
    C
    Return
End Sub

Private Sub ZZ_TakBetBkt()
Dim A$
Ept = "1234()567": A = "sdklfjdsf(1234()567)aaa(": GoSub Tst
Ept = "AA":        A = "XXX(AA)XX":   GoSub Tst
Ept = "A$()A":     A = "(A$()A)XX":   GoSub Tst
Ept = "O$()":      A = "(O$()) As X": GoSub Tst
Exit Sub
Tst:
    Act = TakBetBkt(A)
    C
    Return
End Sub
