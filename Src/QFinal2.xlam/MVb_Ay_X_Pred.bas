Attribute VB_Name = "MVb_Ay_X_Pred"
Option Explicit

Function AyPredAllTrue(A, Pred$) As Boolean
If Sz(A) = 0 Then Exit Function
AyPredAllTrue = ItrPredAllTrue(A, Pred)
End Function

Function AyPredSomFalse(A, Pred$) As Boolean
If Sz(A) = 0 Then Exit Function
AyPredSomFalse = ItrPredSomFalse(A, Pred)
End Function

Sub AyPredSplitAsg(A, Pred$, OTrueAy, OFalseAy)
Dim O1, O2
O1 = AyCln(A)
O2 = O1
Dim X
For Each X In AyNz(A)
    If Run(Pred, X) Then
        Push OTrueAy, X
    Else
        Push OFalseAy, X
    End If
Next
End Sub

Function AyPred_HasSomTrue(A, Pred$) As Boolean
If Sz(A) = 0 Then Exit Function
AyPred_HasSomTrue = ItrPredSomTrue(A, Pred)
End Function

Function AyPred_IsAllFalse(A, Pred$) As Boolean
If Sz(A) = 0 Then Exit Function
AyPred_IsAllFalse = ItrPredAllFalse(A, Pred)
End Function
