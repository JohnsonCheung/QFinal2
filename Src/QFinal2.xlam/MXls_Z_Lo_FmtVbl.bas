Attribute VB_Name = "MXls_Z_Lo_FmtVbl"
Option Explicit

Function QtPrpLoFmlVbl$(A As QueryTable)
QtPrpLoFmlVbl = FbtStrPrpLoFmlVbl(QtFbtStr(A))
End Function

Function QtPrpLoFmtVbl$(A As QueryTable)
If IsNothing(A) Then Exit Function
QtPrpLoFmtVbl = FbtStrPrpLoFmlVbl(QtFbtStr(A))
End Function

Property Get DbtPrpLoFmlVbl$(A As Database, T)
DbtPrpLoFmlVbl = DbtPrp(A, T, "LoFmlVbl")
End Property

Property Let DbtPrpLoFmlVbl(A As Database, T, LoFmlVbl$)
DbtPrp(A, T, "LoFmlVbl") = LoFmlVbl
End Property

Property Get TblPrpLoFmlVbl$(T)
TblPrpLoFmlVbl = DbtPrpLoFmlVbl(CurDb, T)
End Property

Property Let TblPrpLoFmlVbl(T, LoFmlVbl$)
DbtPrpLoFmlVbl(CurDb, T) = LoFmlVbl
End Property

Function LoPrpLoFmlVbl$(A As ListObject)
LoPrpLoFmlVbl = QtPrpLoFmlVbl(LoQt(A))
End Function

Property Get FbtPrpLoFmlVbl$(A$, T$)
FbtPrpLoFmlVbl = DbtPrpLoFmlVbl(FbDb(A), T)
End Property

Property Let FbtPrpLoFmlVbl(A$, T$, LoFmlVbl$)
DbtPrpLoFmlVbl(FbDb(A), T) = LoFmlVbl
End Property

Function FbtStrPrpLoFmlVbl$(FbtStr$)
Dim Fb$, T$
FbtStrAsg FbtStr, Fb, T
FbtStrPrpLoFmlVbl = FbtPrpLoFmlVbl(Fb, T)
End Function

Function WtPrpLoFmlVbl$(T$)
'WtPrpLoFmlVbl = FbtPrpLoFmlVbl(WFb, T)
End Function

Function TpMainPrpLoFmlVbl$()
'TpMainPrpLoFmlVbl = LoPrpLoFmlVbl(TpMainLo)
End Function
