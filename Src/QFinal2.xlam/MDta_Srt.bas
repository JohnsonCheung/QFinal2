Attribute VB_Name = "MDta_Srt"
Option Explicit

Function DrsSrt(A As Drs, ColNm$, Optional IsDes As Boolean) As Drs
Set DrsSrt = Drs(A.Fny, DrySrt(A.Dry, AyIx(A.Fny, ColNm), IsDes))
End Function

Function DrySrt(Dry, ColIx%, Optional IsDes As Boolean) As Variant()
Dim Col: Col = DryCol(Dry, ColIx)
Dim Ix&(): Ix = AySrtIntoIxAy(Col, IsDes)
Dim J%, O()
For J = 0 To UB(Ix)
   Push O, Dry(Ix(J))
Next
DrySrt = O
End Function

Function DtSrt(A As Dt, ColNm$, Optional IsDes As Boolean) As Dt
Set DtSrt = Dt(A.DtNm, A.Fny, DrsSrt(DtDrs(A), ColNm, IsDes).Dry)
End Function
