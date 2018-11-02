Attribute VB_Name = "MVb_Align_Ay"
Option Explicit

Function AyAlignNTerm(A, N%) As String()
Dim W%(), L
W = ZWdtAy(A, N)
For Each L In AyNz(A)
    PushI AyAlignNTerm, AyAlignNTerm1(L, W)
Next
End Function

Function AyAlignT1(A) As String()
Dim T1$(), Rest$()
    AyAsgT1AyRestAy A, T1, Rest
T1 = AyAlignL(T1)
AyAlignT1 = AyabAddWSpc(T1, Rest)
End Function

Private Function AyAlignNTerm1$(A, W%())
Dim Ay$(), J%, N%, O$(), I
N = Sz(W)
Ay = LinNTermRst(A, N)
If Sz(Ay) <> N + 1 Then Stop
For J = 0 To N - 1
    PushI O, AlignL(Ay(J), W(J))
Next
PushI O, Ay(N)
AyAlignNTerm1 = RTrim(JnSpc(O))
End Function

Private Function ZWdtAy(A, NTerm%) As Integer()
If Sz(A) = 0 Then Exit Function
Dim O%(), W%(), L
ReDim O(NTerm - 1)
For Each L In A
    W = ZWdtAy1(L, NTerm)
    O = ZWdtAy2(O, W)
Next
ZWdtAy = O
End Function
Private Function ZWdtAy1(Lin, N%) As Integer()
Dim T
For Each T In LinNTerm(Lin, N)
    PushI ZWdtAy1, Len(T)
Next
End Function
Private Function ZWdtAy2(A%(), B%()) As Integer()
Dim O%(), J%, I
O = A
For Each I In B
    If I > O(J) Then O(J) = I
    J = J + 1
Next
ZWdtAy2 = O
End Function

Function AyAlign1T(A) As String()
AyAlign1T = AyAlignNTerm(A, 1)
End Function

Function AyAlign2T(A) As String()
AyAlign2T = AyAlignNTerm(A, 2)
End Function

Function AyAlign3T(A$()) As String()
AyAlign3T = AyAlignNTerm(A, 3)
End Function

Function AyAlignL(Ay) As String()
Dim W%: W = AyWdt(Ay) + 1
Dim I
For Each I In AyNz(Ay)
    Push AyAlignL, AlignL(I, W)
Next
End Function
Private Sub Z_AyAlign2T()
Dim Ly$()
Ly = ApSy("AAA B C D", "A BBB CCC")
Ept = Sy("AAA B   C D", _
         "A   BBB CCC")
GoSub Tst
Exit Sub
Tst:
    Act = AyAlign2T(Ly)
    C
    Return
End Sub
Private Sub Z_AyAlign3T()
Dim Ly$()
Ly = ApSy("AAA B C D", "A BBB CCC")
Ept = Sy("AAA B   C   D", _
         "A   BBB CCC")
GoSub Tst
Exit Sub
Tst:
    Act = AyAlign3T(Ly)
    C
    Return
End Sub

