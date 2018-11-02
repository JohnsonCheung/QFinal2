Attribute VB_Name = "MVb_X_S1S2"
Option Explicit

Private Function ZZS1S2Ay1() As S1S2()
Dim O() As S1S2
PushObj O, S1S2("sldjflsdkjf", "lksdjf")
PushObj O, S1S2("sldjflsdkjf", "lksdjf")
PushObj O, S1S2("sldjf", "lksdjf")
PushObj O, S1S2("sldjdkjf", "lksdjf")
ZZS1S2Ay1 = O
End Function

Function AyabS1S2Ay(A, B) As S1S2()
If Not AyIsEqSz(A, B) Then Stop
Dim X, J&
For Each X In AyNz(A)
    PushObj AyabS1S2Ay, S1S2(X, B(J))
    J = J + 1
Next
End Function

Function AyMapS1S2Ay(A, MapFunNm$) As S1S2()
Dim O() As S1S2, I
If Sz(A) > 0 Then
    For Each I In A
        PushObj O, S1S2(I, Run(MapFunNm, I))
    Next
End If
AyMapS1S2Ay = O
End Function

Function S1S2Trim(S1, S2, NoTrim) As S1S2
If NoTrim Then
    Set S1S2Trim = S1S2(S1, S2)
Else
    Set S1S2Trim = S1S2(Trim(S1), Trim(S2))
End If
End Function

Function CvS1S2(A) As S1S2
Set CvS1S2 = A
End Function

Function DicS1S2Ay(A As Dictionary) As S1S2()
Dim K
For Each K In A.Keys
    PushObj DicS1S2Ay, S1S2(K, VarLines(A(K)))
Next
End Function

Function S1S2(S1, S2) As S1S2
Set S1S2 = New S1S2
S1S2.S1 = S1
S1S2.S2 = S2
End Function

Sub S1S2Asg(A As S1S2, O1, O2)
O1 = A.S1
O2 = A.S2
End Sub

Function S1S2Clone(A As S1S2) As S1S2
Set S1S2Clone = S1S2(A.S1, A.S2)
End Function

Function S1S2Lin$(A As S1S2, Optional Sep$ = " ", Optional W1%)
S1S2Lin = AlignL(A.S1, W1) & Sep & A.S2
End Function

Function S1S2AyAddAsLy(A() As S1S2, Optional Sep$ = "") As String()
Dim O$(), J&
For J = 0 To UB(A)
   Push O, A(J).S1 & Sep & A(J).S2
Next
S1S2AyAddAsLy = O
End Function

Sub S1S2AyBrw(A() As S1S2)
Brw S1S2AyFmt(A)
End Sub

Function S1S2AyDic(A() As S1S2) As Dictionary
Dim J&, O As New Dictionary
For J = 0 To UB(A)
    With A(J)
        If Not O.Exists(.S1) Then
            O.Add .S1, .S2
        End If
    End With
Next
Set S1S2AyDic = O
End Function

Function S1S2AySy1(A() As S1S2) As String()
Dim O$(), J&
For J = 0 To UB(A)
   Push O, A(J).S1
Next
S1S2AySy1 = O
End Function

Function S1S2AySy2(A() As S1S2) As String()
Dim O$(), J&
For J = 0 To UB(A)
   Push O, A(J).S2
Next
S1S2AySy2 = O
End Function


Function S1S2AySq(A() As S1S2, Optional Nm1$ = "S1", Optional Nm2$ = "S2") As Variant()
If Sz(A) = 0 Then Exit Function
Dim O(), I, R&
ReDim O(1 To Sz(A), 1 To 2)
R = 2
O(1, 1) = Nm1
O(1, 2) = Nm2
For Each I In AyNz(A)
    With CvS1S2(I)
        O(R, 1) = .S1
        O(R, 2) = .S2
        R = R + 1
    End With
Next
S1S2AySq = O
End Function

Function S1S2AyStrDic(A) As Dictionary
Set S1S2AyStrDic = S1S2AyDic(S1S2Ay(A))
End Function

Function S1S2Ay(S1S2AyStr) As S1S2()
Dim I
For Each I In AyNz(Split(S1S2AyStr, "|"))
    PushObj S1S2Ay, BrkBoth(I, ":")
Next
End Function

Function SyPair_S1S2Ay(Sy1$(), Sy2$()) As S1S2()
Ass AyIsEqSz(Sy1, Sy2)
If Sz(Sy1) <> 0 Then Exit Function
Dim U&, O() As S1S2
ReDim O(U)
Dim J&
For J = 0 To UB(Sy1)
    O(J) = S1S2(Sy1(J), Sy2(J))
Next
SyPair_S1S2Ay = O
End Function

Function SyS1S2Ay(A$(), Sep$) As S1S2()
Dim O() As S1S2, J%
Dim U&: U = UB(A)
O = S1S2Ay(U)
For J = 0 To U
    With Brk1(A(J), Sep)
        O(J) = S1S2(.S1, .S2)
    End With
Next
SyS1S2Ay = O
End Function

Private Sub Z_DicS1S2Ay()
Dim A As New Dictionary
A.Add "A", "BB"
A.Add "B", "CCC"
Dim Act() As S1S2
Act = DicS1S2Ay(A)
Stop
End Sub
