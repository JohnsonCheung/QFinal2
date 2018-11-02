Attribute VB_Name = "MDao__Res"
Option Explicit
Sub EnsResTbl()
DbEnsResTbl CurDb
End Sub
Sub DbEnsResTbl(A As Database)
If Not DbHasTbl(A, "Res") Then DbCrtResTbl A
End Sub

Function DbResExp$(A As Database, ResNm)
'Resnm is Tbl.Fld.Key  With Tbl-Dft and Fld-Dft as Res
'Export the res to tmpFfn and return tmpFfn
Dim O$
O = TmpFfn(".txt")
DbResAttFld(A, ResNm).SaveToFile O
DbResExp = O
End Function
