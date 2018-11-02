Attribute VB_Name = "MIde_Gen_Const"
Option Explicit

Function ConstAy_ConstValDry(Cons, A) As Variant()
If Sz(A) = 0 Then Exit Function
Dim O(), I
For Each I In A
   Push O, Array(Cons, I)
Next
ConstAy_ConstValDry = O
End Function

Sub ConstEdt(Nm)
FtBrw FtEns(ConstFt(Nm))
End Sub

Private Function ConstFt$(Nm)
ConstFt = SpecPth & Nm & ".Const.txt"
End Function

Sub ConstGen(Nm)
AyWrt ConstNewLy(Nm), ConstFt(Nm)
ConstEdt Nm
End Sub

Private Function ZConstLy(Nm) As String()
Dim A$, O$()
A = ConstFt(Nm)
If Not FfnIsExist(A) Then Exit Function
Dim F%, L$
F = FtInpHd(A)
While Not EOF(F)
    Line Input #F, L
'    If L = ConstSepLin Then Close F: ZConstLy = O: Exit Function
    Push O, L
Wend
Close #F
ZConstLy = O
End Function

Function ConstNewLy(Nm) As String()
Dim A$(), B$()
A = ZConstLy(Nm)
If Sz(A) = 0 Then Exit Function
B = ZConstLinAy(A, Nm)
'ConstNewLy = AyAddAp(A, Sy(ConstSepLin), B)
End Function
Private Function ZConstLinAy(ConstLy$(), VarNm) As String() 'Ly(A)
Dim N%
N = Sz(ConstLy)
If N <= 20 Then
    ZConstLinAy = ZChunk1(ConstLy, VarNm, "Public")
    Exit Function
End If
Dim VarNy$(), Lin$, Ay$()
VarNy = ZNmNy(VarNm, (N - 1) \ 20 + 1)
Ay = ZChunk(ConstLy, VarNy)
Lin = ZLasLin(VarNm, VarNy)
ZConstLinAy = AyAddItm(Ay, Lin)
End Function
Private Function ZNmNy(Nm, N%) As String()
Dim Fmt$, J%
'Fmt = StrDup("0", N_NDig(N))
For J = 1 To N
    PushI ZNmNy, Nm & "_" & Format(J, Fmt)
Next
End Function
Private Function ZChunk(ConstLy$(), VarNy$()) As String()
Dim J%, Ay$(), O$()
For J = 0 To UB(VarNy)
    Ay = AyMid(ConstLy, J * 20, 20)
    PushAy O, ZChunk1(Ay, VarNy(J), "Private")
Next
ZChunk = O
End Function

Function ZLasLin$(VarNm, Ny$())
Dim B$
B = Join(Ny, " & vbCrLf & ")
ZLasLin = FmtQQ("Const ?$ = ?", VarNm, B)
End Function

Function ZChunk1(A, VarNm, Mdy$) As String()
If Sz(A) = 0 Then
    ZChunk1 = Sy(FmtQQ("? Const ?$ = """"", Mdy, VarNm))
    Exit Function
End If
Dim O$(), L$
Dim J&, U&
U = UB(A)
For J = 0 To U
    L = QuoteAsVb(A(J))
    Select Case True
    Case J = 0
        Push O, FmtQQ("? Const ?$ = ? & _", Mdy, VarNm, L)
    Case J = U
        Push O, "vbCrLf & " & L
    Case Else
        Push O, "vbCrLf & " & L & " & _"
    End Select
Next
ZChunk1 = O
End Function

Function LinesConstLines$(A$, VarNm$)
'LinesConstLines = JnCrLf(ZConstLy(SplitCrLf(A), VarNm))
End Function
