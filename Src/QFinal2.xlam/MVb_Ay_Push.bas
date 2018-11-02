Attribute VB_Name = "MVb_Ay_Push"
Option Explicit

Sub Push(O, M)
Dim N&
N = Sz(O)
ReDim Preserve O(N)
If IsObject(M) Then
    Set O(N) = M
Else
    O(N) = M
End If
End Sub

Sub PushAp(O, ParamArray Ap())
Dim Av(), I: Av = Ap
For Each I In Av
    Push O, I
Next
End Sub

Sub PushAy(O, Ay)
Dim I
For Each I In AyNz(Ay)
    Push O, I
Next
End Sub

Sub PushAyNoDup(O, Ay)
Dim I
For Each I In AyNz(Ay)
    PushNoDup O, I
Next
End Sub

Sub PushDic(O As Dictionary, A As Dictionary)
If IsNothing(O) Then
    Set O = A
    Exit Sub
End If
Dim K
For Each K In A.Keys
    If O.Exists(K) Then Stop
    O.Add K, A(K)
Next
End Sub

Sub PushI(O, M)
Dim N&
N = Sz(O)
ReDim Preserve O(N)
O(N) = M
End Sub

Sub PushIAy(O, MAy)
Dim M
For Each M In AyNz(MAy)
    PushI O, M
Next
End Sub

Sub PushISomSz(OAy, IAy)
If Sz(IAy) = 0 Then Exit Sub
PushI OAy, IAy
End Sub

Sub PushItmAy(O, Itm, Ay)
Push O, Itm
PushAy O, Ay
End Sub

Sub PushNoDup(O, M)
If Not AyHas(O, M) Then PushI O, M
End Sub
Sub PushNoDupNonBlankStr(O, M)
If M = "" Then Exit Sub
If Not AyHas(O, M) Then PushI O, M
End Sub

Sub PushNoDupAy(O, Ay)
Dim I
For Each I In AyNz(Ay)
    PushNoDup O, I
Next
End Sub

Sub PushNonBlankStr(O, M)
If M <> "" Then PushI O, M
End Sub

Sub PushNonBlankSy(O, Sy$())
Dim I
For Each I In AyNz(Sy)
    PushNonBlankStr O, I
Next
End Sub

Sub PushNonEmp(O, M)
If Not IsEmp(M) Then Push O, M
End Sub

Sub PushNonNothing(O, M)
If IsNothing(M) Then Exit Sub
PushObj O, M
End Sub

Sub PushNonZSz(O, Ay)
If Sz(Ay) = 0 Then Exit Sub
PushI O, Ay
End Sub

Sub PushObj(O, M)
Dim N&
N = Sz(O)
ReDim Preserve O(N)
Set O(N) = M
End Sub

Sub PushObjAy(O, Oy)
Dim I
For Each I In AyNz(Oy)
    PushObj O, I
Next
End Sub

Sub PushWithSz(O, Ay)
If Not IsArray(Ay) Then Stop
If Sz(Ay) = 0 Then Exit Sub
Push O, Ay
End Sub

Function Sz&(A)
On Error Resume Next
Sz = UBound(A) + 1
End Function
Function UB&(A)
UB = Sz(A) - 1
End Function


Function Pop(O)
Pop = AyLasEle(O)
If Sz(O) = 1 Then
    Erase O
Else
    ReDim Preserve O(UB(O) - 1)
End If
End Function

Sub ReSz(A, U&)
If U < 0 Then
    Erase A
Else
    ReDim Preserve A(U)
End If
End Sub
