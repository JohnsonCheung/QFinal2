Attribute VB_Name = "MVb_Lin_Lines"
Option Explicit
Private Sub ZZ_LinesAyLyPad()
Dim A$()
Push A, RplVBar("ksdjlfdf|sdklfjdsfdsksdf|skldfjdf")
Push A, RplVBar("ksdjlfdf|sdklfjdsfdsksdf|skldfjdf|sdf")
Push A, RplVBar("ksdjlfdf|sdklfjdsfdf|skldfjdf|lskdf|slkdjf|sdlf||")
Push A, RplVBar("ksdjlfdf|sdklfjsdfdsfdsf|skldsdffjdf")
D LinesAyLyPad(A)
End Sub


Private Sub ZZ_LinesEndTrim()
Dim Lines$: Lines = RplVBar("lksdf|lsdfj|||")
Dim Act$: Act = LinesEndTrim(Lines)
Debug.Print Act & "<"
Stop
End Sub

Sub ZZ_LinesKeepLasN()
Dim Ay$(), A$, J%
For J = 0 To 9
Push Ay, "Line " & J
Next
A = Join(Ay, vbCrLf)
'Debug.Print fLasN(A, 3)
End Sub

Function LinesApp$(A, L)
If A = "" Then LinesApp = L: Exit Function
LinesApp = A & vbCrLf & L
End Function

Function LinesAy_Wdt%(A$())
Dim O%, J&, M%
For J = 0 To UB(A)
   M = LinesWdt(A(J))
   If M > O Then O = M
Next
LinesAy_Wdt = O
End Function

Function LinesAyLyPad(A$()) As String()
LinesAyLyPad = LyPad(LinesAyLy(A))
End Function

Function LinesAyLy(LinesAy) As String()
Dim Lines
For Each Lines In LinesAy
    PushIAy LinesAyLy, SplitLines(CStr(Lines))
Next
End Function

Function LinesAyLines$(A$())
LinesAyLines = JnCrLf(LinesAyLy(A))
End Function

Function LinesAyWdt%(A$())
If Sz(A) = 0 Then Exit Function
Dim O%, J&, M%, L
For Each L In A
   O = Max(O, LinesWdt(CStr(L)))
Next
LinesAyWdt = O
End Function

Function LinesBoxLy(A$) As String()
LinesBoxLy = LyBoxLy(SplitCrLf(A))
End Function

Sub LinesBrkAsg(A$, Ny0, ParamArray OLyAp())
Dim Ny$(), L, T1$, T2$, NmDic As Dictionary
Ny = CvNy(Ny0)
Set NmDic = AyIxDic(Ny)
For Each L In SplitCrLf(A)
    Select Case FstChr(L)
    Case "'", " "
    Case Else
        BrkAsg L, " ", T1, T2
        If NmDic.Exists(T1) Then
            Push OLyAp(NmDic(T1)), T2 '<----
        End If
    End Select
Next
End Sub


Function LinesEndTrim$(A$)
LinesEndTrim = JnCrLf(LyEndTrim(SplitCrLf(A)))
End Function

Sub Z_LinesEndTrim()
Dim Lines$: Lines = RplVBar("lksdf|lsdfj|||")
Dim Act$: Act = LinesEndTrim(Lines)
Debug.Print Act & "<"
Stop
End Sub

Function LinesKeepLasN$(A$, N%)
LinesKeepLasN = JnCrLf(AyKeepLasN(SplitCrLf(A), N))
End Function

Function LinesLasLin$(A$)
If A = "" Then Exit Function
LinesLasLin = AyLasEle(LinesLy(A))
End Function

Function LinesLinCnt&(Lines$)
LinesLinCnt = Sz(SplitCrLf(Lines))
End Function

Function LinesLy(A$) As String()
LinesLy = SplitLines(A)
End Function

Function LinesSplit(A$) As String()
LinesSplit = SplitCrLf(A)
End Function

Function LinesSqH(A$) As Variant()
LinesSqH = AySqH(LinesLy(A))
End Function

Function LinesSqV(A$) As Variant()
LinesSqV = AySqV(LinesLy(A))
End Function

Function LinesTrimEnd$(A$)
LinesTrimEnd = Join(LyTrimEnd(SplitCrLf(A)), vbCrLf)
End Function

Function LinesTab$(Lines$, Optional Space% = 4)
Dim O$(), S$, L
S = VBA.Space(Space)
For Each L In AyNz(SplitCrLf(Lines))
    PushI O, S & L
Next
LinesTab = JnCrLf(O)
End Function

Function LinesVbl$(A$)
Const CSub$ = "LinesVbl"
If HasSubStr(A, "|") Then Er CSub, "Given [Lines] has |", A
LinesVbl = Replace(A, vbCrLf, "|")
End Function

Function LinesWdt%(A)
LinesWdt = AyWdt(SplitLines(A))
End Function
