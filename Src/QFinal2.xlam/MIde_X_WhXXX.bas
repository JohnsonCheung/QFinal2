Attribute VB_Name = "MIde_X_WhXXX"
Option Explicit

Function WhBExprSqp$(BExpr$)
If BExpr = "" Then Exit Function
WhBExprSqp = " Where " & BExpr
End Function

Function WhEmpNm() As WhNm
End Function

Function WhEmpPjMth() As WhPjMth
End Function

Function WhMd(Optional WhCmpTy$, Optional Nm As WhNm) As WhMd
Dim O As New WhMd
O.InCmpTy = CvWhCmpTy(WhCmpTy)
Set O.Nm = Nm
Set WhMd = O
End Function

Function WhMdMth(Optional Md As WhMd, Optional Mth As WhMth) As WhMdMth
Set WhMdMth = New WhMdMth
With WhMdMth
    Set .Md = Md
    Set .Mth = Mth
End With
End Function

Function WhMdMthMd(A As WhMdMth) As WhMd
If Not IsNothing(A) Then Set WhMdMthMd = A.Md
End Function

Function WhMdMthMth(A As WhMdMth) As WhMth
If Not IsNothing(A) Then Set WhMdMthMth = A.Mth
End Function

Function WhMth(Optional WhMdy$, Optional WhKd$, Optional Nm As WhNm) As WhMth
Set WhMth = New WhMth
With WhMth
    .InShtKd = CvWhMthKd(WhKd)
    .InShtMdy = CvWhMdy(WhMdy)
    Set .Nm = Nm
End With
End Function

Function WhPjMth(Optional Pj As WhNm, Optional MdMth As WhMdMth) As WhPjMth
Set WhPjMth = New WhPjMth
With WhPjMth
    Set .Pj = Pj
    Set .MdMth = MdMth
End With
End Function

Function WhPjMth_MdMth(A As WhPjMth) As WhMdMth
If IsNothing(A) Then Exit Function
Set WhPjMth_MdMth = A.MdMth
End Function

Function WhPjMth_Nm(A As WhPjMth) As WhNm
If IsNothing(A) Then Exit Function
Set WhPjMth_Nm = A.Pj
End Function


Function CvWhCmpTy(WhCmpTy$) As vbext_ComponentType()
Dim O() As vbext_ComponentType, I
For Each I In AyNz(SslSy(WhCmpTy))
    Push O, CmpStrTy(I)
Next
CvWhCmpTy = O
End Function

Function CvWhMdy(WhMdy$) As String()
If WhMdy = "" Then Exit Function
Dim O$(), M
O = SslSy(WhMdy): CvWhMdy1 O
If AyHas(O, "Pub") Then Push O, ""
CvWhMdy = O
End Function
Private Function CvWhMdy1(A$())
Dim M
For Each M In A
    If Not AyHas(ShtMdyAy, M) Then Stop
Next
End Function

Function CvWhMthKd(WhMthKd$) As String()
If WhMthKd = "" Then Exit Function
Dim O$(), K
O = SslSy(WhMthKd)
For Each K In O
    If Not AyHas(MthKdAy, K) Then Stop
Next
CvWhMthKd = O
End Function
