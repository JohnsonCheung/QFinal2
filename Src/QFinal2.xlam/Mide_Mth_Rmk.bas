Attribute VB_Name = "Mide_Mth_Rmk"
Option Explicit

Sub MthUnRmk(A As Mth)
Dim P() As FTNo: P = MthCxtFT(A)
Dim J%
For J = UB(P) To 0 Step -1
    MthCxtFT_UnRmk A, P(J)
Next
End Sub



Private Sub ZZ_MthRmk()
Dim M As Mth: Set M = Mth(Md("ZZModule"), "YYA")
            Ass LinesVbl(MthLines(M)) = "Property Get ZZA()|End Property||Property Let YYA(V)||End Property"
MthRmk M:   Ass LinesVbl(MthLines(M)) = "Property Get ZZA()|Stop '|End Property||Property Let YYA(V)|Stop '|'|End Property"
MthUnRmk M: Ass LinesVbl(MthLines(M)) = "Property Get ZZA()|End Property||Property Let YYA(V)||End Property"
End Sub
