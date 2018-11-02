Attribute VB_Name = "MVb_Str_Likss"
Option Explicit
Function StrLikss(A, Likss) As Boolean
StrLikss = StrLikAy(A, SslSy(Likss))
End Function

Function StrLikAy(A, LikAy$()) As Boolean
Dim I
For Each I In AyNz(LikAy)
    If A Like I Then StrLikAy = True: Exit Function
Next
End Function

Function StrLikssAy(A, LikssAy) As Boolean
Dim Likss
For Each Likss In LikssAy
    If StrLikss(A, Likss) Then StrLikssAy = True: Exit Function
Next
End Function


Private Sub Z_T1LikSslAy_T1()
Dim A$(), Nm$
A = SplitVBar("a bb* *dd | c x y")
Nm = "x"
Ept = "c"
GoSub Tst
Exit Sub
Tst:
    Act = T1LikSslAy_T1(A, Nm)
    C
    Return
End Sub


Function T1LikLikSslAy_FstT2T3Eq$(A$(), T2, T3)
Dim L, T$, Lik$, Likssl$
If Sz(A) = 0 Then Exit Function
For Each L In A
    LinAsg2TRst L, T, Lik, Likssl
    If T2 Like Lik Then
        If StrLikss(T3, Likssl) Then
            T1LikLikSslAy_FstT2T3Eq = L
            Exit Function
        End If
    End If
Next
End Function

Function T1LikLikSslAy_T1$(A$(), T2, T3)
Dim L, T$, Lik$, Likssl$
If Sz(A) = 0 Then Exit Function
For Each L In A
    LinAsg2TRst L, T, Lik, Likssl
    If T2 Like Lik Then
        If StrLikss(T3, L) Then
            T1LikLikSslAy_T1 = T
            Exit Function
        End If
    End If
Next
End Function

Function T1LikSslAy_T1$(T1LikSslAy$(), Nm)
Dim L, T1$
If Sz(T1LikSslAy) = 0 Then Exit Function
For Each L In T1LikSslAy
    T1 = ShfT(L)
    If StrLikss(Nm, L) Then
        T1LikSslAy_T1 = T1
        Exit Function
    End If
Next
End Function
