Attribute VB_Name = "MVb_Lin_Scl"
Option Explicit
Sub SclAsg(A$, Ny0, ParamArray Ap())
Dim Av(): Av = Ap
Dim V, Ny$(), I, J%
Ny = CvNy(Ny0)
If Sz(Ny) <> Sz(Av) Then Stop
For Each I In AyNz(AyRmvEmp(AyTrim(SplitSC(A))))
    V = SclItm_V(CStr(I), Ny)
    Select Case True
    Case IsByt(V) And (V = 1 Or V = 2)
    Case IsBool(V) Or IsStr(V): Ap(J) = V
    Case Else: Er "SclAsg", "Program error in SclItm.  It should return (Byt1,Byt2,Bool,Str), but now it returns [Ty]", TypeName(V)
    End Select
    J = J + 1
Next
End Sub

Function SclChk(A$, Ny0) As String()
Dim V, Ny$(), I, Er1$(), Er2$()
Ny = CvNy(Ny0)
For Each I In AyNz(AyRmvEmp(AyTrim(SplitSC(A))))
    V = SclItm_V(CStr(I), Ny)
    Select Case True
    Case IsByt(V) And V = 1: Push Er1, I
    Case IsByt(V) And V = 2: Push Er2, I
    Case IsBool(V) Or IsStr(V)
    Case Else: Er "SclChk", "Program error in SclItm.  It should return (Byt1,Byt2,Bool,Str), but now it returns [Ty]", TypeName(V)
    End Select
Next
Dim O$()
    If Sz(Er1) > 0 Then
        O = MsgLy("There are [invalid-SclNy] in given [scl] under these [valid-SclNy].", JnSpc(Er1), A, JnSpc(Ny))
    End If
    If Sz(Er2) > 0 Then
        PushAy O, MsgLy("[Itm] of [Scl] has [valid-SclNy], but it is not one of SclNy nor it has '='", Er2, A, Ny)
    End If
SclChk = O
End Function

Function SclItm_V(A$, Ny$())
'Return Byt1 if Pfx of A not in Ny
'Return True If A = One Of Ny
'Return Byt2 if Pfx of A is in Ny, but not Eq one Ny and Don't have =
If AyHas(Ny, A) Then SclItm_V = True: Exit Function
If Not StrMatchPfxAy(A, Ny) Then SclItm_V = CByte(1): Exit Function
If Not HasSubStr(A, "=") Then SclItm_V = CByte(2): Exit Function
SclItm_V = Trim(TakAft(A, "="))
End Function

Function SclShf$(OA)
BrkS1Asg OA, ";", SclShf, OA
End Function
