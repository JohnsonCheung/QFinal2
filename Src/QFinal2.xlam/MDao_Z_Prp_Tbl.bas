Attribute VB_Name = "MDao_Z_Prp_Tbl"
Private Sub ZZ_DbtPrp()
TblDrp "Tmp"
DoCmd.RunSQL "Create Table Tmp (F1 Text)"
DbtPrp(CurDb, "Tmp", "XX") = "AFdf"
Debug.Assert DbtPrp(CurDb, "Tmp", "XX") = "AFdf"
End Sub

Function DbtCrtPrp(A As Database, T, P$, V) As DAO.Property
Set DbtCrtPrp = A.TableDefs(T).CreateProperty(P, VarDaoTy(V), V)
End Function

Function DbtHasPrp(A As Database, T, P$) As Boolean
DbtHasPrp = ItrHasNm(A.TableDefs(T).Properties, P)
End Function
Property Get DbtPrp(A As Database, T, P$)
If Not DbtHasPrp(A, T, P) Then Exit Property
DbtPrp = A.TableDefs(T).Properties(P).Value
End Property

Property Let DbtPrp(A As Database, T, P$, V)
If DbtHasPrp(A, T, P) Then
    A.TableDefs(T).Properties(P).Value = V
Else
    A.TableDefs(T).Properties.Append DbtCrtPrp(A, T, P, V)
End If
End Property
