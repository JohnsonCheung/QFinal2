Attribute VB_Name = "MDao_CurDb_Tf"
Option Explicit
Sub Z()
End Sub
Function TfEleScl$(T, F)
TfEleScl = DbtfEleScl(CurDb, T, F)
End Function

Property Get TfidV(T$, F$, Id&)
TfidV = DbtfidV(CurDb, T, F, Id)
End Property

Property Let TfidV(T$, F$, Id&, V)
DbtfidV(CurDb, T$, F$, Id) = V
End Property

Property Get Tfk0V(T$, F$, K0)
Tfk0V = Dbtfk0V(CurDb, T, F, K0)
End Property

Property Let Tfk0V(T$, F$, K0, V)
Dbtfk0V(CurDb, T, F, K0) = V
End Property

Function TfkV(T$, F$, ParamArray K())
Dim K0(): K0 = K
TfkV = Dbtfk0V(CurDb, T, F, K0)
End Function


Function TfTy(T, F) As DAO.DataTypeEnum
TfTy = DbtfTy(CurDb, T, F)
End Function
