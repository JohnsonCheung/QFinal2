Attribute VB_Name = "MTp_Tp_Lin_Cln"
Option Explicit
Function ClnBrk1(A$(), Ny0) As Variant()
Dim O(), U%, Ny$(), L, T1$, T2$, NmDic As Dictionary, Ix%, Er$()
Ny = CvNy(Ny0)
U = UB(Ny)
ReDim O(U)
O = AyMap(O, "EmpSy")
Set NmDic = AyIxDic(Ny)
For Each L In A
    LinAsgTRst LTrim(L), T1, T2
    If NmDic.Exists(T1) Then
        Ix = NmDic(T1)
        Push O(Ix), T2 '<----
    End If
Next
Push O, ClnT1Chk(A, Ny)
ClnBrk1 = O
End Function

Function ClnT1Chk(A$(), T1Ay0) As String()
Dim T1Ay$(), L, O$()
T1Ay = CvNy(T1Ay0)
For Each L In A
    If Not AyHas(T1Ay, LinT1(L)) Then Push O, L
Next
If Sz(O) > 0 Then
    O = AyAddPfx(AyQuoteSqBkt(O), Space(4))
    O = AyIns(O, FmtQQ("Following lines have invalid T1.  Valid T1 are [?]", JnSpc(T1Ay)))
End If
ClnT1Chk = O
End Function

Function LinCln$(A)
If IsEmp(A) Then Exit Function
If LinIsDotLin(A) Then Exit Function
If LinIsSngTerm(A) Then Exit Function
If LinIsDDLin(A) Then Exit Function
LinCln = TakBefDD(A)
End Function

Function LyCln(A) As String()
LyCln = AyRmvEmp(AyMapSy(A, "LinCln"))
End Function

Function LyClnLnxAy(A) As Lnx()
Dim O()  As Lnx, L$, J%
For J = 0 To UB(A)
    L = LinCln(A(J))
    If L <> "" Then
        Dim M  As Lnx
        Set M = New Lnx
        M.Ix = J
        M.Lin = A(J)
        Push O, M
    End If
Next
LyClnLnxAy = O
End Function
