Attribute VB_Name = "MVb_X_LnoCnt"
Option Explicit
Function FmCnt(FmLno, Cnt) As FmCnt
Dim O As New FmCnt
Set FmCnt = O.Init(FmLno, Cnt)
End Function

Function CvFmCnt(A) As FmCnt
Set CvFmCnt = A
End Function

Function FmCntStr$(A As FmCnt)
FmCntStr = "FmLno[" & A.FmLno & "] Cnt[" & A.Cnt & "]"
End Function
Function FmCntAyIsInOrd(A() As FmCnt) As Boolean
Dim J%
For J = 0 To UB(A) - 1
    With A(J)
      If .FmLno + .Cnt > A(J + 1).FmLno Then Exit Function
    End With
Next
FmCntAyIsInOrd = True
End Function

Function FmCntAyIsEq(A() As FmCnt, B() As FmCnt) As Boolean
If Sz(A) <> Sz(B) Then Exit Function
Dim X, J&
For Each X In AyNz(A)
    If Not FmCntIsEq(CvFmCnt(X), B(J)) Then Exit Function
    J = J + 1
Next
FmCntAyIsEq = True
End Function
Function FmCntIsEq(A As FmCnt, B As FmCnt) As Boolean
With A
    If .FmLno <> B.FmLno Then Exit Function
    If .Cnt <> B.Cnt Then Exit Function
End With
FmCntIsEq = True
End Function

Function FmCntAyLinCnt%(A() As FmCnt)
Dim I, C%, O%
For Each I In A
    C = CvFmCnt(I).Cnt
    If C > 0 Then O = O + C
Next
FmCntAyLinCnt = O
End Function

Function LnoCnt(Lno&, Cnt&) As LnoCnt
Dim O As New LnoCnt
Set LnoCnt = O.Init(Lno, Cnt)
End Function
