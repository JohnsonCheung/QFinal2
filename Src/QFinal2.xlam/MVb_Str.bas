Attribute VB_Name = "MVb_Str"
Option Explicit

Function AddLbl$(A, Lbl$)
Dim B$
If IsDate(A) Then
    B = DteDTim(A)
Else
    B = Replace(Replace(A, ";", "%3B"), "=", "%3D")
End If
If A <> "" Then AddLbl = Lbl & "=" & B
End Function

Function StrIsEq(A, B) As Boolean
StrIsEq = StrComp(A, B, vbBinaryCompare) = 0
End Function

Function StrApp$(A, L)
If A = "" Then StrApp = L: Exit Function
StrApp = A & " " & L
End Function
Function Pad0$(N, NDig%)
Pad0 = Format(N, StrDup("0", NDig))
End Function

Function StrBrk1(A, Sep$, Optional NoTrim As Boolean) As String()
StrBrk1 = StrBrk1At(A, InStr(A, Sep), Sep, NoTrim)
End Function

Function StrBrk1At(A, At&, Sep, Optional NoTrim As Boolean) As String()
Dim O1$, O2$
If At = 0 Then
    O1 = A
Else
    O1 = Left(A, At - 1)
    O2 = Mid(A, At + Len(Sep))
End If
If Not NoTrim Then
    O1 = Trim(O1)
    O2 = Trim(O2)
End If
StrBrk1At = ApSy(O1, O2)
End Function

Function StrBrk(A, Sep$, Optional NoTrim As Boolean) As String()
StrBrk = StrBrkAt(A, InStr(A, Sep), Sep, NoTrim)
End Function

Function StrBrkAt(A, At&, Sep, Optional NoTrim As Boolean) As String()
If At = 0 Then Stop
Dim O1$, O2$
O1 = Left(A, At - 1)
O2 = Mid(A, At + Len(Sep))
If Not NoTrim Then
    O1 = Trim(O1)
    O2 = Trim(O2)
End If
StrBrkAt = ApSy(O1, O2)
End Function

Sub StrBrw(A, Optional Fnn$)
Dim T$: T = TmpFt("StrBrw", Fnn$)
StrWrt A, T
FtBrw T
End Sub

Function StrDft$(A, B)
StrDft = IIf(A = "", B, A)
End Function

Function StrDup$(S, N%)
Dim O$, J%
For J = 0 To N - 1
    O = O & S
Next
StrDup = O
End Function

Function StrInSfxAy(A, SfxAy$()) As Boolean
StrInSfxAy = AyHasPredPXTrue(SfxAy, "HasSfx", A)
End Function

Function StrMatchPfxAy(A, PfxAy$()) As Boolean
If Sz(PfxAy) = 0 Then Exit Function
Dim Pfx
For Each Pfx In PfxAy
    If A Like Pfx & "*" Then StrMatchPfxAy = True: Exit Function
Next
End Function

Sub StrWrt(A, Ft, Optional OvrWrt As Boolean)
If OvrWrt Then FfnDltIfExist (Ft)
Fso.CreateTextFile(Ft, True).Write A
End Sub

Function SubStrCnt&(A, SubStr)
Dim P&: P = 1
Dim L%: L = Len(SubStr)
Dim O%
While P > 0
    P = InStr(P, A, SubStr)
    If P = 0 Then SubStrCnt = O: Exit Function
    O = O + 1
    P = P + L
Wend
SubStrCnt = O
End Function

Function SubStrPos(A, SubStr$) As FTIx
Dim FmIx&: FmIx = InStr(A, SubStr)
Dim ToIx&
If FmIx > 0 Then ToIx = FmIx + Len(SubStr)
SubStrPos = FTIx(FmIx, ToIx)
End Function

Sub Z()
Z_SubStrCnt
End Sub

Sub Z_SubStrCnt()
Dim A$, SubStr$
A = "aaaa":                 SubStr = "aa":  Ept = CLng(2): GoSub Tst
A = "aaaa":                 SubStr = "a":   Ept = CLng(4): GoSub Tst
A = "skfdj skldfskldf df ": SubStr = " ":   Ept = CLng(3): GoSub Tst
Exit Sub
Tst:
    Act = SubStrCnt(A, SubStr)
    C
    Return
End Sub
