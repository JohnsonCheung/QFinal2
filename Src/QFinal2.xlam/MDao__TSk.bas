Attribute VB_Name = "MDao__TSk"
Option Explicit
Function TSkIsExist(T, ParamArray SkAp()) As Boolean
Dim Sk(): Sk = SkAp
TSkIsExist = DbtSkAvIsExist(CurDb, T, Sk)
End Function

Sub TkIns(T, ParamArray K())
Dim K0(): K0 = K
Dbtk0Ins CurDb, T, K0
End Sub

Function DbtSkAvIsExist(A As Database, T, SkAv()) As Boolean
DbtSkAvIsExist = DbqAny(A, QSel_Fm(T, DbtSkAvWhStr(A, T, SkAv)))
End Function

Function TsfV(T, S, F) ' S is Ssk-Value  (Ssk is single-field-secondary-key)
TsfV = DbtsfV(CurDb, T, S, F)
End Function
Function DbtsfV(A As Database, T, S, F) ' S is Ssk-Value (Ssk is single-field-secondary-key)
Dim W$
W = X.Whfv(DbtSsk(A, T), S)
DbtsfV = DbqV(A, QSel_FF_Fm(F, T, W))
End Function
