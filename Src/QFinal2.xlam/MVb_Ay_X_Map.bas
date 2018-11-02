Attribute VB_Name = "MVb_Ay_X_Map"
Option Explicit

Function AyMap(A, Map$)
AyMap = AyMapInto(A, Map, EmpAy)
End Function

Function AyMapABXInto(Ay, ABX$, A, B, OInto)
Dim O: O = OInto
Erase O
If Sz(Ay) > 0 Then
    Dim J&, X
    ReDim O(UB(A))
    For Each X In AyNz(A)
        Asg Run(ABX, A, B, X), O(J)
        J = J + 1
    Next
End If
AyMapABXInto = O
End Function

Function AyMapABXSy(Ay, ABX$, A, B) As String()
AyMapABXSy = AyMapABXInto(Ay, ABX, A, B, EmpSy)
End Function

Function AyMapAXBInto(Ay, AXB$, A, B, OInto)
Dim O: O = OInto
Erase O
If Sz(Ay) > 0 Then
    Dim J&, X
    ReDim O(UB(A))
    For Each X In AyNz(A)
        Asg Run(AXB, A, X, B), O(J)
        J = J + 1
    Next
End If
AyMapAXBInto = O
End Function

Function AyMapAXBSy(Ay, AXB$, A, B)
AyMapAXBSy = AyMapAXBInto(Ay, AXB, A, B, EmpSy)
End Function

Function AyMapAsgAy(A, OAy, MthNm$, ParamArray Ap())
If Sz(A) = 0 Then Exit Function
Dim Av(): Av = Ap
Av = AyIns(Av)
Dim I, J&
Dim O
O = OAy
Erase O
Dim U&: U = UB(A)
    ReDim O(U)
For Each I In A
    Asg I, Av(0)
    Asg RunAv(MthNm, Av), O(J)
    J = J + 1
Next
AyMapAsgAy = O
End Function

Function AyMapAsgSy(A, MthNm$, ParamArray Ap()) As String()
If Sz(A) = 0 Then Exit Function
Dim Av(): Av = Ap
If Sz(Av) = 0 Then
    AyMapAsgSy = AyMapSy(A, MthNm)
    Exit Function
End If
Dim I, J&
Dim O$()
    ReDim O(UB(A))
    Av = AyIns(Av)
    For Each I In A
        Asg I, Av(0)
        Asg RunAv(MthNm, Av), O(J)
        J = J + 1
    Next
AyMapAsgSy = O
End Function

Function AyMapAvInto(A, MapFunNm$, PrmAv, OIntoAy)
Dim O: O = OIntoAy: Erase O
If Sz(A) > 0 Then
    Dim I
    Stop
    Dim Av(): Av = PrmAv: Av = AyIns(PrmAv)
    For Each I In A
        Asg I, Av(0)
        Push O, RunAv(MapFunNm, Av)
    Next
End If
AyMapAvInto = O
End Function

Function AyMapAvSy(A, MapFunNm$, PrmAv) As String()
AyMapAvSy = AyMapAvInto(A, MapFunNm, PrmAv, EmpSy)
End Function

Sub AyMapDo(A, Map$, DoFun$)
Dim X
For Each X In AyNz(A)
    Run DoFun, Run(Map, X)
Next
End Sub

Function AyMapFlat(Ay, ItmAyFun$)
AyMapFlat = AyMapFlatInto(Ay, ItmAyFun, EmpAy)
End Function

Function AyMapFlatInto(Ay, ItmAyFun$, OIntoAy)
Dim O, J&, M
O = OIntoAy: Erase O
For J = 0 To UB(Ay)
    M = Run(ItmAyFun, Ay(J))
    PushAy O, M
Next
AyMapFlatInto = O
End Function

Function AyMapInto(A, MapFunNm$, OIntoAy)
Dim O: O = OIntoAy: Erase OIntoAy
Dim I
If Sz(A) > 0 Then
    For Each I In A
        Push O, Run(MapFunNm, I)
    Next
End If
AyMapInto = O
End Function

Function AyMapLngAy(A, MapMthNm$) As Long()
AyMapLngAy = AyMapInto(A, MapMthNm, EmpLngAy)
End Function

Function AyMapObjFunXInto(A, Obj, FunX$, OInto)
Dim O, X
O = OInto
Erase O
If Sz(A) > 0 Then
    For Each X In AyNz(A)
        Push O, CallByName(Obj, FunX, VbMethod, X)
    Next
End If
AyMapObjFunXInto = O
End Function

Function AyMapPX(A, XP$, P) As Variant()
AyMapPX = AyMapPXInto(A, XP, P, EmpAy)
End Function

Function AyMapPXInto(A, PX$, P, OInto)
'MapPXFunNm cannot be Excel-like-Function-Name, eg A2, AA2, (cell-address)
' or "Mth"
AyMapPXInto = AyCln(OInto)
Dim X
For Each X In AyNz(A)
    Push AyMapPXInto, Run(PX, P, X)
Next
End Function

Function AyMapPXSy(A, PX$, P) As String()
AyMapPXSy = AyMapPXInto(A, PX, P, EmpSy)
End Function

Function AyMapSy(A, MapMthNm$) As String()
AyMapSy = AyMapInto(A, MapMthNm, EmpSy)
End Function

Function AyMapXAB(Ay, XAB$, A, B)
AyMapXAB = AyMapXABInto(Ay, XAB, A, B, EmpSy)
End Function

Function AyMapXABCDInto(Ay, XABC$, A, B, C, D, OInto)
Erase OInto
If Sz(Ay) = 0 Then AyMapXABCDInto = OInto: Exit Function
Dim X
For Each X In AyNz(A)
    Push OInto, Run(XABC, X, A, B, C, D)
Next
AyMapXABCDInto = OInto
End Function

Function AyMapXABCInto(Ay, XABC$, A, B, C, OInto)
Erase OInto
Dim X
For Each X In AyNz(A)
    Push OInto, Run(XABC, X, A, B, C)
Next
AyMapXABCInto = OInto
End Function

Function AyMapXABInto(Ay, XAB$, A, B, OInto)
'MapXPFunNm cannot be Excel-like-Function-Name, eg A2, AA2, (cell-address)
AyMapXABInto = AyCln(OInto)
Dim X
For Each X In AyNz(Ay)
    Push AyMapXABInto, Run(XAB, X, A, B)
Next
End Function

Function AyMapXABSy(Ay, XAB$, A, B) As String()
AyMapXABSy = AyMapXABInto(Ay, XAB, A, B, EmpSy)
End Function

Function AyMapXAP(A, MthNm$, ParamArray Ap()) As Variant()
If Sz(A) = 0 Then Exit Function
Dim Av(): Av = Ap
Av = AyIns(Av)
Dim I, J&
Dim O()
Dim U&: U = UB(A)
    ReDim O(U)
For Each I In A
    Asg I, Av(0)
    Asg RunAv(MthNm, Av), O(J)
    J = J + 1
Next
AyMapXAP = O
End Function

Function AyMapXP(A, XP$, P) As Variant()
AyMapXP = AyMapXPInto(A, XP, P, EmpAy)
End Function

Function AyMapXPInto(A, XP$, P, OInto)
'MapXPFunNm cannot be Excel-like-Function-Name, eg A2, AA2, (cell-address)
Dim O, X
O = OInto
Erase O
For Each X In AyNz(A)
    Push O, Run(XP, X, P)
Next
AyMapXPInto = O
End Function

Function AyMapXPSy(A, XP$, P) As String()
AyMapXPSy = AyMapXPInto(A, XP, P, EmpSy)
End Function

Private Sub ZZ_AyMap()
Dim Act: Act = AyMap(Array(1, 2, 3, 4), "Mul2")
Ass Sz(Act) = 4
Ass Act(0) = 2
Ass Act(1) = 4
Ass Act(2) = 6
Ass Act(3) = 8
End Sub

Private Sub ZZ_AyMapSy()
Dim Ay$(): Ay = AyMapSy(Array("skldfjdf", "aa"), "RmvFstChr")
Stop
End Sub

Sub Z_AyMapSy()
Dim Ay$(): Ay = AyMapSy(Array("skldfjdf", "aa"), "RmvFstChr")
Stop
End Sub

Private Sub Z_AyMapXAP()
Dim Act: Act = AyMap(Array(1, 2, 3, 4), "Mul2")
Ass Sz(Act) = 4
Ass Act(0) = 2
Ass Act(1) = 4
Ass Act(2) = 6
Ass Act(3) = 8
End Sub
