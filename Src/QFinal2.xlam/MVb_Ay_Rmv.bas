Attribute VB_Name = "MVb_Ay_Rmv"
Option Explicit
Function AyRmv3T(A) As String()
AyRmv3T = AyMapSy(A, "Rmv3T")
End Function

Function AyRmvDDLin(A) As String()
AyRmvDDLin = AyWhPredFalse(A, "LinIsDDLin")
End Function

Function AyRmvDotLin(A) As String()
AyRmvDotLin = AyWhPredFalse(A, "LinIsDotLin")
End Function

Function AyRmvEle(A, Ele)
Dim Ix&: Ix = AyIx(A, Ele): If Ix = -1 Then AyRmvEle = A: Exit Function
AyRmvEle = AyRmvEleAt(A, AyIx(A, Ele))
End Function

Function AyRmvNEle(A, Ele, Cnt%)
If Cnt <= 0 Then Stop
AyRmvNEle = AyCln(A)
Dim X, C%
C = Cnt
For Each X In AyNz(A)
    If C = 0 Then
        PushI AyRmvNEle, X
    Else
        If X <> Ele Then
            Push AyRmvNEle, X
        Else
            C = C - 1
        End If
    End If
Next
X:
End Function

Function AyRmvEleAt(Ay, Optional At = 0, Optional Cnt = 1)
AyRmvEleAt = AyWhExlAtCnt(Ay, At, Cnt)
End Function

Function AyRmvEleLik(A, Lik$) As String()
If Sz(A) = 0 Then Exit Function
Dim J&
For J = 0 To UB(A)
    If A(J) Like Lik Then AyRmvEleLik = AyRmvEleAt(A, J): Exit Function
Next
End Function

Function AyRmvEmp(A)
Dim O: O = AyCln(A)
If Sz(A) > 0 Then
    Dim X
    For Each X In AyNz(A)
        PushNonEmp O, X
    Next
End If
AyRmvEmp = O
End Function

Function AyRmvEmpEleAtEnd(A)
Dim LasU&, U&
Dim O: O = A
For LasU = UB(A) To 0 Step -1
    If Not IsEmp(O(LasU)) Then
        Exit For
    End If
Next
If LasU = -1 Then
    Erase O
Else
    ReDim Preserve O(LasU)
End If
AyRmvEmpEleAtEnd = O
End Function
Function AyRmvFmTo(A, FmIx, ToIx)
Dim U&
U = UB(A)
If 0 > FmIx Or FmIx > U Then Stop
If ToIx > FmIx Or FmIx > U Then Stop
Dim O
    O = A
    Dim I&, J&
    I = 0
    For J = ToIx + 1 To U
        O(FmIx + I) = O(J)
        I = I + 1
    Next
    Dim Cnt&
    Cnt = ToIx - FmIx + 1
    ReDim Preserve O(U - Cnt)
AyRmvFmTo = O
End Function
Function AyRmvFTIx(A, B As FTIx)
With B
    AyRmvFTIx = AyRmvFmTo(A, .FmIx, .ToIx)
End With
End Function

Function AyRmvFstChr(A) As String()
AyRmvFstChr = AyMapSy(A, "RmvFstChr")
End Function

Function AyRmvFstEle(A)
AyRmvFstEle = AyRmvEleAt(A)
End Function

Function AyRmvFstNEle(A, N)
Dim O: O = A
ReDim O(N - 1)
Dim J&
For J = 0 To UB(A) - N
    O(J) = A(N + J)
Next
AyRmvFstNEle = O
End Function

Function AyRmvFstNonLetter(A) As String()
AyRmvFstNonLetter = AyMapSy(A, "RmvFstNonLetter")
End Function

Function AyRmvLasChr(A) As String()
AyRmvLasChr = AyMapSy(A, "RmvLasChr")
End Function

Function AyRmvLasEle(A)
AyRmvLasEle = AyRmvEleAt(A, UB(A))
End Function

Function AyRmvLasNEle(A, Optional NEle% = 1)
Dim O: O = A
Select Case Sz(A)
Case Is > NEle:    ReDim Preserve O(UB(A) - NEle)
Case NEle: Erase O
Case Else: Stop
End Select
AyRmvLasNEle = O
End Function

Function AyRmvOneTermLin(A) As String()
AyRmvOneTermLin = AyWhPredFalse(A, "LinIsOneTermLin")
End Function

Function AyRmvSngQRmk(A$()) As String()
If Sz(A) = 0 Then Exit Function
Dim X, O$()
For Each X In AyNz(A)
    If Not IsSngQRmk(CStr(X)) Then Push O, X
Next
AyRmvSngQRmk = O
End Function

Function AyRmvSngQuote(A$()) As String()
AyRmvSngQuote = AyMapSy(A, "RmvSngQuote")
End Function

Function AyRmvT1(A) As String()
Dim I
For Each I In AyNz(A)
    PushI AyRmvT1, RmvT1(I)
Next
End Function

Function AyRmvTT(A$()) As String()
AyRmvTT = AyMapSy(A, "RmvTT")
End Function

Private Sub Z_AyRmvFTIx()
Dim A
Dim FTIx1 As FTIx
Dim Act
A = SplitSpc("a b c d e")
Set FTIx1 = FTIx(1, 2)
Act = AyRmvFTIx(A, FTIx1)
Ass Sz(Act) = 3
Ass JnSpc(Act) = "a d e"
End Sub


Private Sub Z_AyRmvEmpEleAtEnd()
Dim A: A = Array(Empty, Empty, Empty, 1, Empty, Empty)
Dim Act: Act = AyRmvEmpEleAtEnd(A)
Ass Sz(Act) = 4
Ass Act(3) = 1
End Sub

Private Sub Z_AyRmvFTIx1()
Dim A
Dim Act
A = SplitSpc("a b c d e")
Act = AyRmvFTIx(A, FTIx(1, 2))
Ass Sz(Act) = 3
Ass JnSpc(Act) = "a d e"
End Sub

Function AyRmvNeg(A)
Dim I
AyRmvNeg = AyCln(A)
For Each I In AyNz(A)
    If I >= 0 Then
        PushI AyRmvNeg, I
    End If
Next
End Function


Function AyRmvPfx(A, Pfx) As String()
If Sz(A) = 0 Then Exit Function
Dim U&: U = UB(A)
Dim O$()
ReDim O(U)
Dim J&
For J = 0 To U
    O(J) = RmvPfx(A(J), Pfx)
Next
AyRmvPfx = O
End Function
