Attribute VB_Name = "MVb_Ay_Fst"
Option Explicit
Option Compare Text
Function AyFstEle(Ay)
If Sz(Ay) = 0 Then Exit Function
Asg Ay(0), AyFstEle
End Function

Function AyFstEqV(A, V)
If AyHas(A, V) Then AyFstEqV = V
End Function

Function AyFstLik$(A, Lik$)
If Sz(A) = 0 Then Exit Function
Dim X
For Each X In A
    If X Like Lik Then AyFstLik = X: Exit Function
Next
End Function

Function AyFstLikItm(A, Lik, Itm)
AyFstLikItm = AyFstPredXABYes(A, "LinHasLikItm", Lik, Itm)
End Function

Function AyFstNEle(A, N)
Dim O: O = A
ReDim Preserve O(N - 1)
AyFstNEle = O
End Function

Function AyFstPfx$(PfxAy, Lin$)
Dim P
For Each P In PfxAy
    If HasPfx(Lin, CStr(P)) Then AyFstPfx = P: Exit Function
Next
End Function

Function AyFstPredPX(A, PX$, P)
If Sz(A) = 0 Then Exit Function
Dim X
For Each X In AyNz(A)
    If Run(PX, P, X) Then Asg X, AyFstPredPX: Exit Function
Next
End Function

Function AyFstPredXABYes(Ay, XAB$, A, B)
Dim X
For Each X In AyNz(Ay)
    If Run(XAB, X, A, B) Then Asg X, AyFstPredXABYes: Exit Function
Next
End Function

Function AyFstPredXP(A, XP$, P)
If Sz(A) = 0 Then Exit Function
Dim X
For Each X In AyNz(A)
    If Run(XP, X, P) Then Asg X, AyFstPredXP: Exit Function
Next
End Function

Function AyFstT1$(A, T1, Optional CasSen As Boolean)
Dim L
For Each L In AyNz(A)
    If LinT1(L) = T1 Then AyFstT1 = L: Exit Function
Next
End Function

Function AyFstRmvT1$(A, T1)
AyFstRmvT1 = RmvT1(AyFstT1(A, T1))
End Function

Function AyFstT2$(A, T2)
AyFstT2 = AyFstPredXP(A, "LinHasT2", T2)
End Function

Function AyFstTT$(A, T1, T2)
AyFstTT = AyFstPredXABYes(A, "LinHasTT", T1, T2)
End Function
