Attribute VB_Name = "MVb_Lin_Vbl"
Option Explicit
Function Vbl_LasLin$(Vbl)
Vbl_LasLin = AyLasEle(SplitVBar(Vbl))
End Function

Function Vbl_Wdt%(Vbl$)
Ass IsVdtVbl(Vbl)
Vbl_Wdt = AyWdt(SplitVBar(Vbl))
End Function



Function VblAy_IsVdt(A$()) As Boolean
If Sz(A) = 0 Then VblAy_IsVdt = True: Exit Function
Dim I
For Each I In A
    If Not IsVdtVbl(CStr(I)) Then Exit Function
Next
VblAy_IsVdt = True
End Function

Function VblAyWdt%(VblAy$())
Dim W%(), J%
For J = 0 To UB(VblAy)
    Push W, VblWdt(VblAy(J))
Next
VblAyWdt = AyMax(W)
End Function

Function VblDic(Vbl, Optional JnSep$ = vbCrLf) As Dictionary
Set VblDic = LyDic(SplitVBar(Vbl), JnSep)
End Function

Function VblLasLin$(Vbl)
VblLasLin = AyLasEle(SplitVBar(Vbl))
End Function

Function VblLines$(Vbl, Optional Pfx$, Optional Ident0%, Optional Sfx$, Optional Wdt0%)
VblLines = JnCrLf(VblLy(VblLines(Vbl), Pfx, Ident0, Sfx, Wdt0))
End Function

Function VblLy(Vbl$, Optional Pfx$, Optional Ident0%, Optional Sfx$, Optional Wdt0%) As String()
Ass IsVdtVbl(Vbl)
If Vbl = "" Then Exit Function
Dim Wdt%
    Wdt = Vbl_Wdt(Vbl)
    If Wdt < Wdt0 Then
        Wdt = Wdt0
    End If
Dim Ident%
    If Ident < 0 Then
        Ident = 0
    Else
        Ident = Ident0
    End If
    If Pfx <> "" Then
        If Ident < Len(Pfx) Then
            Ident = Len(Pfx) + 1
        End If
    End If
Dim O$()
    Dim Ay$()
    Ay = SplitVBar(Vbl)
    Dim J%, A$, U&, S$, S1$, P$
    U = UB(Ay)
    P = IIf(Pfx <> "", Pfx & " ", "")
    S1 = Space(Ident)
    For J = 0 To U
        If J = 0 Then
'            S = AlignL(P, Ident, DoNotCut:=True)
        Else
            S = S1
        End If
'        A = S & AlignL(Ay(J), Wdt, ErIfNotEnoughWdt:=True)
        If J = U Then
            A = A & " " & Sfx
        End If
        Push O, A
    Next
VblLy = O
End Function

Function VblLyDry(A$()) As Variant()
If Sz(A) = 0 Then Exit Function
Dim O()
   Dim I
   For Each I In A
       Push O, AyTrim(SplitVBar(CStr(I)))
   Next
VblLyDry = O
End Function

Sub Z_VblLyDry()
Dim VblLy$()
Push VblLy, "|lskdf|sdlf|lsdkf"
Push VblLy, "|lsdf|"
Push VblLy, "|lskdfj|sdlfk|sdlkfj|sdklf|skldf|"
Push VblLy, "|sdf"
GoSub Tst
Exit Sub
Tst:
    Act = VblLyDry(VblLy)
    Ass DryIsEq(CvAy(Act), CvAy(Ept))
    Return
End Sub
Function VblWdt%(Vbl$)
Ass IsVbl(Vbl)
VblWdt = AyWdt(VblLy(Vbl))
End Function

Private Sub ZZ_Vbl_Wdt()
Dim Act%: Act = Vbl_Wdt("lksdjf|sldkf|              df")
Ass Act = 16
End Sub

Private Sub ZZ_VblLy()
AyDmp VblLy("lksfj|lksdfjldf|lskdlksdflsdf|sdkjf", "Select")
End Sub

Private Sub ZZ_VblLyDry()
Dim VblLy$()
Dim Exp$()
Push VblLy, "|lskdf|sdlf|lsdkf"
Push VblLy, "|lsdf|"
Push VblLy, "|lskdfj|sdlfk|sdlkfj|sdklf|skldf|"
Push VblLy, "|sdf"
Dim Act()
Act = VblLyDry(VblLy)
End Sub
