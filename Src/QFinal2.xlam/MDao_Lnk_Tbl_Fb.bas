Attribute VB_Name = "MDao_Lnk_Tbl_Fb"
Option Explicit
Sub DbLnkFb(A As Database, Fb$, Tny0, Optional SrcTny0)
Dim Tny$(): Tny = CvNy(Tny0)              ' Src_Tny
Dim Src$(): Src = CvNy(Dft(SrcTny0, Tny0)) ' Tar_Tny
Ass Sz(Tny) > 0
Ass Sz(Src) = Sz(Tny)
Dim J%
For J = 0 To UB(Tny)
    DbtLnkFb A, Tny(J), Fb$, Src(J)
Next
End Sub

Function DbLnkTny(A As Database) As String()
DbLnkTny = ItrWhPredPrpSy(A.TableDefs, "TdHasCnStr", "Name")
End Function


Function DbtLnkFb(A As Database, T, Fb$, Optional Fbt0$) As String()
Dim Fbt$, Cn$
Cn = ";Database=" & Fb
Fbt = DftStr(Fbt0, T)
DbtLnkFb = DbtLnk(A, T, Fbt, Cn)
End Function
