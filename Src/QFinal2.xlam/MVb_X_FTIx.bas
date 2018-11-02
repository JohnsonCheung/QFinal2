Attribute VB_Name = "MVb_X_FTIx"
Option Explicit
Function FTIxAyFC(A() As FTIx) As FmCnt()
FTIxAyFC = AyMapInto(A, "FTIxFC", FTIxAyFC)
End Function

Function FTIxAyLnoCntAy(A() As FTIx) As LnoCnt()
Dim U&, J&
    U = UB(A)
Dim O() As LnoCnt
   ReDim O(U)
For J = 0 To U
   O(J) = FTIxLnoCnt(A(J))
Next
FTIxAyLnoCntAy = O
End Function

Function FTIxIsEmp(A As FTIx) As Boolean
FTIxIsEmp = True
If A.FmIx < 0 Then Exit Function
If A.ToIx < 0 Then Exit Function
If A.FmIx > A.ToIx Then Exit Function
FTIxIsEmp = False
End Function

Function FTIxCnt&(A As FTIx)
If FTIxIsVdt(A) Then Exit Function
FTIxCnt = A.ToIx - A.FmIx + 1
End Function

Function FTIxFC(A As FTIx) As FmCnt
With A
    Set FTIxFC = FmCnt(.FmIx + 1, .ToIx - .FmIx + 1)
End With
End Function

Function FTIxHasU(A As FTIx, U&) As Boolean
If U < 0 Then Stop
If FTIxIsEmp(A) Then Exit Function
If A.FmIx > U Then Exit Function
If A.ToIx > U Then Exit Function
FTIxHasU = True
End Function

Function FTIxIsVdt(A As FTIx) As Boolean
FTIxIsVdt = True
If A.FmIx < 0 Then Exit Function
If A.ToIx < 0 Then Exit Function
If A.FmIx > A.ToIx Then Exit Function
FTIxIsVdt = False
End Function

Sub FmIxToIxAss(FmIx, ToIx, U)
Const CSub$ = "FmIxToIxAss"
If FmIx < 0 Then Er CSub, "[FmIx] is negative, where [U] and [ToIx]", FmIx, U, ToIx
If ToIx < 0 Then Er CSub, "[ToIx] is negative, where [U] and [FmIx]", ToIx, U, FmIx
End Sub

Function FTIxLinCnt%(A As FTIx)
Dim O%
O = A.ToIx - A.FmIx + 1
If O < 0 Then Stop
FTIxLinCnt = O
End Function

Function FTIxLnoCnt(A As FTIx) As LnoCnt
Dim Lno&, Cnt&
   Cnt = A.ToIx - A.FmIx + 1
   If Cnt < 0 Then Cnt = 0
   Lno = A.FmIx + 1
Set FTIxLnoCnt = LnoCnt(Lno, Cnt)
End Function

Function FTIxNo(A As FTIx) As FTNo
Set FTIxNo = FTNo(A.FmIx + 1, A.ToIx + 1)
End Function

Function FTIx(FmIx, ToIx) As FTIx
Dim O As New FTIx
Set FTIx = O.Init(FmIx, ToIx)
End Function
