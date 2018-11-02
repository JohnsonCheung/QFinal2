Attribute VB_Name = "MDao__Tmp"
Option Explicit
Function TmpTd() As DAO.TableDef
Dim O() As DAO.Field2
Push O, NewFd("F1")
Set TmpTd = NewTd("Tmp", O)
End Function

Function TmpDb(Optional Fdr$, Optional Fnn$) As Database
Dim Fb$
Fb = TmpFb(Fdr, Fnn)
FbCrt Fb
Set TmpDb = FbDb(Fb)
End Function
