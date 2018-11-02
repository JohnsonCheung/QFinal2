Attribute VB_Name = "MVb_Ay__Oy"
Option Explicit
Function Oy_Cat_AyPrp_AsAy(A, AyPrpNm$)
Dim O, J&, I
If Sz(A) = 0 Then Exit Function
O = CallByName(A(0), AyPrpNm, VbGet)
If Not IsArray(O) Then ErPm ' Given AyPrpNm is not of a array-property
For J = 1 To UB(A)  ' from start Ix=1
    I = CallByName(A(J), AyPrpNm, VbGet)
    If Not IsArray(I) Then ErDta
    PushAy O, I
Next
Oy_Cat_AyPrp_AsAy = O
End Function

Function Oy_Map_ByObjGet(A, Obj, GetMthNm$, OIntoAy)
Dim O: O = OIntoAy
Erase O
Dim ArgAy(0), J%
For J = 0 To UB(A)
    Asg A(J), ArgAy(0)
    Push O, CallByName(Obj, GetMthNm, VbGet, ArgAy)
Next
Oy_Map_ByObjGet = O
End Function

Function OyCompoundPrpSy(A, PrpSsl$) As String()
Dim O$(), I
If Sz(A) = 0 Then Exit Function
For Each I In A
    Push O, ObjCompoundPrp(A, PrpSsl)
Next
OyCompoundPrpSy = O
End Function

Sub OyDo(Oy, DoFun$)
Dim O
For Each O In Oy
    Excel.Run DoFun, O ' DoFunNm cannot be like a Excel.Address (eg, A1, XX1)
Next
End Sub

Sub OyDoMth(A, Mth$)
Dim J&
For J = 0 To UB(A)
    CallByName A(J), Mth, VbMethod
Next
End Sub

Sub OyEachSubP1(A, SubNm$, Prm)
If Sz(A) = 0 Then Exit Sub
Dim O
For Each O In A
    CallByName O, SubNm, VbMethod, Prm
Next
End Sub

Function OyFstPrpEqV(A, P, V)
If Sz(A) = 0 Then Exit Function
Dim X
For Each X In AyNz(A)
    If ObjPrp(X, P) = V Then Asg X, OyFstPrpEqV: Exit Function
Next
End Function

Function OyHas(A, Obj) As Boolean
Dim X, Op&
Op = ObjPtr(Obj)
For Each X In AyNz(A)
    If ObjPtr(X) = Op Then OyHas = True: Exit Function
Next
End Function

Function OyMap(A, MapMthNm$) As Variant()
OyMap = OyMapInto(A, MapMthNm, EmpAy)
End Function

Function OyMapInto(A, MapFunNm$, OIntoAy)
Dim Obj, J&, U&
U = UB(A)
Dim O
O = OIntoAy
ReSz O, U
For J = 0 To U
    Asg Run(MapFunNm, A(J)), O(J)
Next
OyMapInto = O
End Function

Function OyNy(A) As String()
OyNy = OyPrpSy(A, "Name")
End Function

Function OyPrpAy(Oy, PrpNm) As Variant()
OyPrpAy = OyPrpAyInto(Oy, PrpNm, EmpAy)
End Function

Function OyPrpAyInto(Oy, PrpNm, OIntoAy)
Dim O: O = OIntoAy: Erase O
If Sz(Oy) > 0 Then
    Dim I
    For Each I In Oy
        Push O, ObjPrp(I, PrpNm)
    Next
End If
OyPrpAyInto = O
End Function

Function OyPrpIntAy(A, PrpNm$) As Integer()
OyPrpIntAy = OyPrpInto(A, PrpNm, EmpIntAy)
End Function

Function OyPrpInto(A, PrpNm$, OInto)
If Sz(A) = 0 Then Exit Function
OyPrpInto = ItrPrpInto(A, PrpNm, OInto)
End Function

Function OyPrpSrtedUniqAy(A, PrpNm$) As Variant()
OyPrpSrtedUniqAy = AySrt(AyWhDist(OyPrpAy(A, PrpNm)))
End Function

Function OyPrpSrtedUniqIntAy(A, PrpNm$) As Integer()
OyPrpSrtedUniqIntAy = AySrt(AyWhDist(OyPrpIntAy(A, PrpNm)))
End Function

Function OyPrpSrtedUniqSy(A, PrpNm$) As Variant()
OyPrpSrtedUniqSy = AySrt(AyWhDist(OyPrpSy(A, PrpNm)))
End Function

Function OyPrpSy(A, PrpNm$)
OyPrpSy = OyPrpInto(A, PrpNm, EmpSy)
End Function

Function OyRmvFstNEle(A, N&)
Dim O: O = A
ReDim O(N - 1)
Dim J&
For J = 0 To UB(A) - N
    Set O(J) = A(N + J)
Next
OyRmvFstNEle = O
End Function

Function OyRmvNothing(A)
OyRmvNothing = AyCln(A)
Dim I
For Each I In A
    If Not IsNothing(I) Then PushObj OyRmvNothing, I
Next
End Function

Function OySrt_By_CompoundPrp(A, PrpSsl$)
Dim O: O = A: Erase O
Dim Sy$(): Sy = OyCompoundPrpSy(A, PrpSsl)
Dim Ix&(): Ix = AySrtIntoIxAy(Sy)
Dim J&
For J = 0 To UB(Ix)
    PushObj O, A(Ix(J))
Next
OySrt_By_CompoundPrp = O
End Function

Function OyToStr$(A)
Dim O$(), I
For Each I In A
    Push O, CallByName(I, "ToStr", VbGet)
Next
OyToStr = JnCrLf(O)
End Function

Function OyWhIxAy(A, IxAy)
Dim O: O = A: Erase O
Dim U&: U = UB(IxAy)
Dim J&
ReSz O, U
For J = 0 To U
    Asg A(IxAy(J)), O(J)
Next
OyWhIxAy = O
End Function

Function OyWhIxSelIntPrp(A, WhIx, PrpNm$) As Integer()
OyWhIxSelIntPrp = OyWhIxSelPrp(A, WhIx, PrpNm, EmpIntAy)
End Function

Function OyWhIxSelPrp(A, WhIx, PrpNm$, OupAy)
Dim Oy1: Oy1 = OyWhIxAy(A, WhIx)  ' Oy1 is subset of Oy
OyWhIxSelPrp = OyPrpInto(Oy1, PrpNm, OupAy)
End Function

Function OyWhIxSelSyPrp(A, WhIx, PrpNm$) As String()
OyWhIxSelSyPrp = OyWhIxSelPrp(A, WhIx, PrpNm, EmpSy)
End Function

Function OyWhNm(A, B As WhNm)
Dim X
For Each X In AyNz(A)
    If IsNmSel(X.Name, B) Then PushObj OyWhNm, X
Next
End Function

Function OyWhNmExl(A, ExlAy$)
If ExlAy = "" Then OyWhNmExl = A: Exit Function
Dim X, LikAy$(), O
O = A
Erase O
LikAy = SslSy(ExlAy)
For Each X In AyNz(A)
    If Not IsInLikAy(X.Name, LikAy) Then PushObj O, X
Next
OyWhNmExl = O
End Function

Function OyWhNmHasPfx(A, Pfx$)
OyWhNmHasPfx = OyWhPredXP(A, "ObjHasNmPfx", Pfx)
End Function

Function OyWhNmPatn(A, Patn$)
If Patn = "." Then OyWhNmPatn = A: Exit Function
Dim X, O, Re As New RegExp
O = A
Erase O
Re.Pattern = Patn
For Each X In AyNz(A)
    If Re.Test(X.Name) Then PushObj O, X
Next
OyWhNmPatn = O
End Function

Function OyWhNmPatnExl(A, Patn$, ExlAy$)
OyWhNmPatnExl = OyWhNmExl(OyWhNmPatn(A, Patn), ExlAy)
End Function

Function OyWhNmReExl(A, Re As RegExp, ExlLikAy$())
If Sz(A) = 0 Then OyWhNmReExl = A: Exit Function
Dim X
For Each X In AyNz(A)
    If IsNmSelReExl(X.Name, Re, ExlLikAy) Then PushObj OyWhNmReExl, X
Next
End Function

Function OyWhPredXP(A, XP$, P)
Dim O, X
O = A
Erase O
For Each X In AyNz(A)
    If Run(XP, X, P) Then
        PushObj A, X
    End If
Next
OyWhPredXP = O
End Function

Function OyWhPrp(A, PrpNm$, PrpEqToVal)
Dim O
   O = A
   Erase O
If Not Sz(A) > 0 Then
   Dim I
   For Each I In A
       If CallByName(I, PrpNm, VbGet) = PrpEqToVal Then PushObj O, I
   Next
End If
End Function

Function OyWhPrpEqV(A, P, V)
Dim X
If Sz(A) > 0 Then
    For Each X In AyNz(A)
        If ObjPrp(X, P) = V Then
            Set OyWhPrpEqV = X: Exit Function
        End If
    Next
End If
Set OyWhPrpEqV = Nothing
End Function

Function OyWhPrpEqValSelPrpInt(A, WhPrpNm$, EqVal, SelPrpNm$) As Integer()
Dim Oy1: Oy1 = OyWhPrpEqV(A, WhPrpNm, EqVal)
OyWhPrpEqValSelPrpInt = OyPrpIntAy(Oy1, SelPrpNm)
End Function

Function OyWhPrpEqValSelPrpSy(A, WhPrpNm$, EqVal, SelPrpNm$) As String()
Dim Oy1: Oy1 = OyWhPrpEqV(A, WhPrpNm, EqVal)
OyWhPrpEqValSelPrpSy = OyPrpSy(Oy1, SelPrpNm)
End Function

Function OyWhPrpIn(A, P, InAy)
Dim X, O
If Sz(A) = 0 Or Sz(InAy) Then OyWhPrpIn = A: Exit Function
O = A
Erase O
For Each X In AyNz(A)
    If AyHas(InAy, ObjPrp(X, P)) Then PushObj O, X
Next
OyWhPrpIn = O
End Function

Function SelOy(A, PrpSsl$) As Variant()

End Function

Sub ZZ_OyDrs()
'WsVis DrsNewWs(OyDrs(CurrentDb.TableDefs("ZZ_DbtUpdSeq").Fields, "Name Type OrdinalPosition"))
End Sub

Private Sub ZZ_OyPrpAy()
Dim CdPanAy() As CodePane
Stop
'CdPanAy = Oy(CurPjx.MdAy).PrpAy("CodePane", CdPanAy)
Stop
End Sub
