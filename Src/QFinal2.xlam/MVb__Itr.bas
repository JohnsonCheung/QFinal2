Attribute VB_Name = "MVb__Itr"
Option Explicit
Function ItrAy(A)
ItrAy = ItrAyInto(A, Array())
End Function

Function ItrAyInto(A, OInto)
ItrAyInto = AyCln(OInto)
Dim X
For Each X In A
    Push ItrAyInto, X
Next
End Function

Function ItrCntTruePrp&(A, BoolPrpNm)
Dim O&, X
For Each X In A
    If CallByName(X, BoolPrpNm, VbGet) Then
        O = O + 1
    End If
Next
ItrCntTruePrp = O
End Function

Sub ItrDo(A, DoFun$)
Dim I
For Each I In A
    Run DoFun, I
Next
End Sub

Sub ItrDoMth(A, MthNm$)
Dim I
For Each I In A
    CallByName I, MthNm, VbMethod
Next
End Sub

Sub ItrDoPX(A, PX$, P)
Dim X
For Each X In A
    Run PX, P, X
Next
End Sub

Sub ItrDoXP(A, XP$, P)
Dim X
For Each X In A
    Run XP, X, P
Next
End Sub

Function ItrFst(A)
Dim X
For Each X In AyNz(A)
    Asg X, ItrFst
    Exit Function
Next
End Function

Function ItrFstItm(A)
Dim X
For Each X In A
    Set ItrFstItm = X
Next
End Function

Function ItrFstNm(A, Nm) ' Return first element in Itr-A with name eq Nm
Dim X
For Each X In A
    If ObjNm(X) = Nm Then Set ItrFstNm = X: Exit Function
Next
Set ItrFstNm = Nothing
End Function

Function ItrFstPredXP(A, XP$, P)
Dim X
For Each X In A
    If Run(XP, X, P) Then Asg X, ItrFstPredXP: Exit Function
Next
End Function

Function ItrFstPrpEqV(A, P, V) 'Return first element in Itr-A with its Prp-P eq to V
Dim X
For Each X In A
    If ObjPrp(X, P) = V Then Set ItrFstPrpEqV = X: Exit Function
Next
Set ItrFstPrpEqV = Nothing
End Function

Function ItrFstPrpTrue(A, P) 'Return first element in Itr-A with its Prp-P being true
Dim X
For Each X In A
    If ObjPrp(X, P) Then Set ItrFstPrpTrue = X: Exit Function
Next
Set ItrFstPrpTrue = Nothing
End Function

Function ItrHas(A, M) As Boolean
Dim I
For Each I In A
    If I = M Then ItrHas = True: Exit Function
Next
End Function

Function ItrHasNm(A, Nm) As Boolean
Dim I
For Each I In A
    If I.Name = Nm Then ItrHasNm = True: Exit Function
Next
End Function

Function ItrHasNmWhRe(A, Re As RegExp) As Boolean
Dim I
For Each I In A
    If Re.Test(I.Name) Then ItrHasNmWhRe = True: Exit Function
Next
End Function

Function ItrHasPrpEqV(A, P, V) As Boolean
Dim X
For Each X In A
    If ObjPrp(X, P) = V Then ItrHasPrpEqV = True: Exit Function
Next
End Function

Function ItrHasPrpTrue(A, P) As Boolean
Dim X
For Each X In A
    If ObjPrp(X, P) Then ItrHasPrpTrue = True: Exit Function
Next
End Function

Function ItrInto(A, OInto)
ItrInto = AyCln(OInto)
Dim X
For Each X In A
    Push ItrInto, X
Next
End Function

Function ItrMap(A, Map$) As Variant()
ItrMap = ItrMapInto(A, Map, EmpAy)
End Function

Function ItrMapInto(A, Map$, OInto)
Dim O: O = OInto
Erase O
Dim X
For Each X In A
    Push O, Run(Map, X)
Next
ItrMapInto = O
End Function

Function ItrMapSy(A, Map$) As String()
ItrMapSy = ItrMapInto(A, Map, EmpSy)
End Function

Function ItrMaxPrp(A, P)
Dim X, O
For Each X In A
    O = Max(O, ObjPrp(X, P))
Next
ItrMaxPrp = O
End Function

Function ItrNy(A) As String()
Dim I
For Each I In A
    PushI ItrNy, ObjNm(I)
Next
End Function

Function ItrNyWhLik(A, Lik) As String()
ItrNyWhLik = AyWhLik(ItrNy(A), Lik)
End Function

Function ItrNyWhPatnExl(A, Optional Patn$, Optional Exl$) As String()
ItrNyWhPatnExl = AyWhPatnExl(ItrNy(A), Patn, Exl)
End Function

Function ItrpAyInto(A, P, OInto)
Dim X, O
O = OInto
Erase O
For Each X In A
    Push O, ObjPrp(X, P)
Next
ItrpAyInto = 0
End Function

Function ItrPredAllFalse(A, Pred$) As Boolean
Dim I
For Each I In A
    If Run(Pred, I) Then Exit Function
Next
ItrPredAllFalse = True
End Function

Function ItrPredAllTrue(A, Pred$) As Boolean
Dim I
For Each I In A
    If Not Run(Pred, I) Then Exit Function
Next
ItrPredAllTrue = True
End Function

Function ItrPredSomFalse(A, Pred$) As Boolean
Dim I
For Each I In A
    If Not Run(Pred, I) Then ItrPredSomFalse = True: Exit Function
Next
End Function

Function ItrPredSomTrue(A, Pred$) As Boolean
Dim I
For Each I In A
    If Run(Pred, I) Then ItrPredSomTrue = True: Exit Function
Next
End Function

Function ItrPrpAy(A, P) As Variant()
ItrPrpAy = ItrPrpAyInto(A, P, EmpAy)
End Function

Function ItrPrpAyInto(A, P, OInto)
ItrPrpAyInto = AyCln(OInto)
Dim I
For Each I In A
    Push ItrPrpAyInto, ObjPrp(I, P)
Next
End Function

Function ItrPrpInto(A, P, OInto)
ItrPrpInto = ItrPrpAyInto(A, P, OInto)
End Function

Function ItrPrpSy(A, P) As String()
ItrPrpSy = ItrPrpInto(A, P, EmpSy)
End Function

Function ItrVy(A) As Variant()
ItrVy = ItrPrpAy(A, "Value")
End Function

Function ItrWhInNyInto(A, InNy$(), OInto)
Dim O, X
O = OInto
Erase O
For Each X In A
    If AyHas(InNy, X.Name) Then PushObj O, X
Next
ItrWhInNyInto = O
End Function

Function ItrWhNm(A, B As WhNm)
ItrWhNm = ItrWhNmInto(A, B, EmpAy)
End Function

Function ItrWhNmInto(A, B As WhNm, OInto)
ItrWhNmInto = AyCln(OInto)
Dim X
For Each X In A
    If IsNmSel(X.Name, B) Then PushObj ItrWhNmInto, X
Next
End Function

Function ItrWhNmReExl(A, Re As RegExp, ExlLikAy$())
ItrWhNmReExl = ItrWhNmReExlInto(A, Re, ExlLikAy, EmpAy)
End Function

Function ItrWhNmReExlInto(A, Re As RegExp, ExlAy$(), OInto)
Dim X
ItrWhNmReExlInto = AyCln(OInto)
For Each X In A
    If IsNmSelReExl(X.Name, Re, ExlAy) Then PushObj ItrWhNmReExlInto, X
Next
End Function

Function ItrWhNyInto(A, InNy$(), OInto)
Dim O, X
O = OInto
Erase O
For Each X In A
    If AyHas(InNy, X.Name) Then PushObj X, O
Next
ItrWhNyInto = O
End Function

Function ItrWhPredPrpAy(A, Pred$, P)
ItrWhPredPrpAy = ItrWhPredPrpAyInto(A, Pred, P, EmpAy)
End Function

Function ItrWhPredPrpAyInto(A, Pred$, P, OInto)
Dim O: O = OInto
Erase O
Dim X
For Each X In A
    If Run(Pred, X) Then
        Push O, ObjPrp(X, P)
    End If
Next
ItrWhPredPrpAyInto = O
End Function

Function ItrWhPredPrpSy(A, Pred$, P) As String()
ItrWhPredPrpSy = ItrWhPredPrpAyInto(A, Pred, P, EmpSy)
End Function

Function ItrWhPrpIsTrue(A, P)
ItrWhPrpIsTrue = ItrWhPrpIsTrueInto(A, P, EmpAy)
End Function
Function ItrClnAy(A)
If A.Count = 0 Then Exit Function
Dim X
For Each X In A
    ItrClnAy = Array(X)
    Exit Function
Next
End Function

Function ItrWhPrpEqV(A, P, V)
Dim O: O = ItrClnAy(A): If IsEmpty(O) Then Exit Function
Dim X
For Each X In A
    If ObjPrp(X, P) = V Then PushObj O, X
Next
ItrWhPrpEqV = O
End Function

Function ItrWhPrpIsTrueInto(A, P, OInto)
Dim O: O = OInto: Erase O
Dim X
For Each X In A
    If ObjPrp(A, P) Then
        Push O, X
    End If
Next
ItrWhPrpIsTrueInto = O
End Function

Function ItrWhWhNm(A, B As WhNm)
ItrWhWhNm = ItrWhNmInto(A, B, EmpAy)
End Function
