Attribute VB_Name = "MIde_Srt"
Option Explicit
Sub Z()
Z_SrcSrtDic
End Sub

Private Sub ZZ_Dcl_BefAndAft_Srt()
Const MdNm$ = "VbStrRe"
Dim A$() ' Src
Dim B$() ' Src->Srt
Dim A1$() 'Src->Dcl
Dim B1$() 'Src->Src->Dcl
A = MdSrc(Md(MdNm))
B = SrcSrtLy(A)
A1 = SrcDclLy(A)
B1 = SrcDclLy(B)
Stop
End Sub

Private Sub Z_SrcSrtDic()

End Sub
Function SrcSrtDic(A$()) As Dictionary
Dim D As Dictionary, K
Set D = SrcDic(A)
Dim O As New Dictionary
    For Each K In D
        O.Add MthDDNmSrtKey(K), D(K)
    Next
Set SrcSrtDic = DicSrt(O)
End Function
Function SrcSrtLines$(A$())
Dim D As Dictionary
Set D = SrcSrtDic(A)
Dim K, O$(), Fst As Boolean
Fst = True
For Each K In AyNz(D.Keys)
    If Fst Then
        Fst = False
        PushI O, D(K)
    Else
        PushI O, vbCrLf & D(K)
    End If
Next
SrcSrtLines = JnCrLf(O)
End Function
Function SrcSrtLy(A$()) As String()
SrcSrtLy = SplitCrLf(SrcSrtLines(A))
End Function

Function CurMdSrtLines$()
CurMdSrtLines = MdSrtLines(CurMd)
End Function
Private Sub ZZ_SrcSrtLy()
Brw SrcSrtLines(CurSrc)
Stop
End Sub





Sub PjSrt(A As VBProject)
Dim M As CodeModule, I, Ay() As CodeModule
If Sz(Ay) = 0 Then Exit Sub
For Each I In Ay
    MdSrt CvMd(I)
Next
End Sub

Function MthNm3SrtKey$(ShtMdy$, ShtTy$, Nm$)
Dim P% 'Priority
    Select Case True
    Case HasPfx(Nm, "Init"): P = 1
    Case Nm = "Z":     P = 9
    Case Nm = "ZZ":    P = 9
    Case HasPfx(Nm, "Z_"):   P = 9
    Case HasPfx(Nm, "ZZ_"):  P = 8
    Case HasPfx(Nm, "Z"):    P = 7
    Case Else:              P = 2
    End Select
MthNm3SrtKey = P & ":" & Nm & ":" & ShtTy & ":" & ShtMdy
End Function

Private Sub ZZ_MthDDNmSrtKey()
GoSub X0
GoSub X1
Exit Sub
X0:
    Dim Ay1$(): Ay1 = SrcMthNy(CurSrc)
    Dim Ay2$(): Ay2 = AyMapSy(Ay1, "MthNmSrtKey")
    S1S2AyBrw AyabS1S2Ay(Ay2, Ay1)
    Return
X1:
    Const A$ = "YYA.Fun."
    Debug.Print MthDDNmSrtKey(A)
    Return
End Sub
Function MthDDNmSrtKey$(A) ' MthDDNm is Nm.ShtTy.ShtMdy
If A = "*Dcl" Then MthDDNmSrtKey = "*Dcl": Exit Function
Dim B$(): B = SplitDot(A): If Sz(B) <> 3 Then Stop
Dim ShtMdy$, ShtTy$, Nm$
AyAsg B, Nm, ShtTy, ShtMdy
MthDDNmSrtKey = MthNm3SrtKey(ShtMdy, ShtTy, Nm)
End Function

Function LinMthSrtKey$(A)
Dim L$, Mdy$, Ty$, Nm$
L = A
Mdy = ShfMdy(L)
Ty = ShfMthTy(L): If Ty = "" Then Exit Function
Nm = TakNm(L)
LinMthSrtKey = MthNm3SrtKey(Mdy, Ty, Nm)
End Function

Private Sub ZZ_LinMthSrtKey()
Dim Ay1$(): Ay1 = SrcMthDclAy(CurSrc)
Dim Ay2$(): Ay2 = AyMapSy(Ay1, "LinMthKey")
S1S2AyBrw AyabS1S2Ay(Ay2, Ay1)
End Sub

Private Sub ZZ_LinMthSrtKey_1()
Const A$ = "Function YYA()"
Debug.Print LinMthSrtKey(A)
End Sub


Function MdSrtLines$(A As CodeModule)
MdSrtLines = SrcSrtLines(MdSrc(A))
End Function

Function MdSrtLy(A As CodeModule) As String()
MdSrtLy = SrcSrtLy(MdSrc(A))
End Function
