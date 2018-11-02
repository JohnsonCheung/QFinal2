Attribute VB_Name = "MDao_Z_Prp"
Option Explicit
Function PrpHas(A As DAO.Properties, P) As Boolean
PrpHas = ItrHasNm(A, P)
End Function

Function PrpVal(A As DAO.Properties, PrpNm$)
On Error Resume Next
PrpVal = A(PrpNm).Value
End Function


Property Let DbtfPrp(A As Database, T, F, P, V)
'If IsEmpty(V) Then
'    If DbtfHasPrp(A, T, F, P) Then
'        A.TableDefs(T).Fields(T).Properties.Delete P
'    End If
'    Exit Function
'End If
If DbtfHasPrp(A, T, F, P) Then
    A.TableDefs(T).Fields(F).Properties(P).Value = V
Else
    With A.TableDefs(T)
        .Fields(F).Properties.Append .CreateProperty(P, VarDaoTy(V), V)
    End With
End If
End Property

Property Get DbtfPrp(A As Database, T, F, P)
If Not DbtfHasPrp(A, T, F, P) Then Exit Property
DbtfPrp = A.TableDefs(T).Fields(F).Properties(P).Value
End Property

Function DbtfHasPrp(A As Database, T, F, P) As Boolean
DbtfHasPrp = ItrHasNm(A.TableDefs(T).Fields(F).Properties, P)
End Function

Property Let TblPrp(T, P$, V)
DbtPrp(CurDb, T, P) = V
End Property
