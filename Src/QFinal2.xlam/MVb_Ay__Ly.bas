Attribute VB_Name = "MVb_Ay__Ly"
Option Explicit
Function LyAy_Lin1$(A(), WdtAy%(), Ix%)
Dim J%, W%, I$, Ly$(), Dr$()
For J = 0 To UB(A)
    Ly = A(J)
    W% = WdtAy(J)
    If UB(Ly) >= Ix Then
        I = Ly(Ix)
    Else
        I = ""
    End If
    Push Dr, AlignL(I, W)
Next
LyAy_Lin1 = "| " + Join(Dr, " | ") + " |"
End Function

Function LyBoxLy(A) As String()
If Sz(A) = 0 Then Exit Function
Dim W%: W = AyWdt(A)
Dim H$: H = "|" & StrDup("-", W + 2) & "|"
Dim O$()
Push O, H
Dim I
For Each I In A
    Push O, "| " & AlignL(I, W) + " |"
Next
Push O, H
LyBoxLy = O
End Function

Function LyEndTrim(A$()) As String()
If Sz(A) = 0 Then Exit Function
If AyLasEle(A) <> "" Then LyEndTrim = A: Exit Function
Dim J%
For J = UB(A) To 0 Step -1
    If Trim(A(J)) <> "" Then
        Dim O$()
        O = A
        ReDim Preserve O(J)
        LyEndTrim = O
        Exit Function
    End If
Next
End Function


Function LyPad(A$()) As String()
Dim W%: W = AyWdt(A)
Dim L
For Each L In AyNz(A)
    PushI LyPad, AlignL(L, W)
Next
End Function

Function LyRmv2Dash(A$()) As String()
If Sz(A) = 0 Then Exit Function
Dim O$(), I
For Each I In A
    Push O, Rmv2Dash(CStr(I))
Next
LyRmv2Dash = O
End Function

Function LySqH(A$()) As Variant()
LySqH = AySqH(A)
End Function

Function LySqV(A$()) As Variant()
LySqV = AySqV(A)
End Function

Function LyT1Ay(A) As String()
Dim O$(), L, J&
If Sz(A) = 0 Then Exit Function
ReDim O(UB(A))
For Each L In A
    BrkAsg L, " ", O(J)
    J = J + 1
Next
End Function

Function LyTrimEnd(Ly) As String()
If Sz(Ly) = 0 Then Exit Function
Dim L$
Dim J&
For J = UB(Ly) To 0 Step -1
    L = Trim(Ly(J))
    If Trim(Ly(J)) <> "" Then
        Dim O$()
        O = Ly
        ReDim Preserve O(J)
        LyTrimEnd = O
        Exit Function
    End If
Next
End Function

Function LyTRst_Dic(A$()) As Dictionary
Dim L, K$, Rst$, O As New Dictionary
For Each L In AyNz(A)
    LinAsgTRst L, K, Rst
    O.Add K, Rst
Next
Set LyTRst_Dic = O
End Function

Function LyWhT1Er(A, Ny0) As String()
'return subset of Ly for those T1 not in Ny0
Dim O$(), T1$, L, Ny$(), U&, NmDic As Dictionary
Ny = CvNy(Ny0)
U = UB(Ny)
Set NmDic = AyIxDic(Ny)
For Each L In A
    T1 = LinT1(L)
    If Not NmDic.Exists(T1) Then
        Push O, L
    End If
Next
LyWhT1Er = O
End Function
