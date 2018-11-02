Attribute VB_Name = "MVb_Dic_LinesDic"
Option Explicit

Sub LinesDic_Brw(A As Dictionary)
AyBrw LinesDicLy(A)
End Sub

Function LinesDicLy(A As Dictionary) As String()
Dim K
For Each K In A.Keys
    PushIAy LinesDicLy, LinesDicLy1(K, A(K))
Next
End Function

Private Function LinesDicLy1(K, Lines) As String()
Dim L
For Each L In AyNz(SplitCrLf(Lines))
    Push LinesDicLy1, K & " " & L
Next
End Function

Function LinesDicLy_LinesDic(A$()) As Dictionary
Dim O As New Dictionary
    Dim L, T1$
    For Each L In AyNz(A)
        T1 = ShfTerm(L)
        If O.Exists(T1) Then
            O(T1) = O(T1) & vbCrLf & L
        Else
            O(T1) = L
        End If
    Next
Set LinesDicLy_LinesDic = O
End Function
