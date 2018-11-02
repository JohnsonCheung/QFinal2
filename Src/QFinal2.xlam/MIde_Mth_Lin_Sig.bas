Attribute VB_Name = "MIde_Mth_Lin_Sig"
Option Explicit

Function MthSig(MthLin$) As MthSig
Dim O As MthSig
With O
Stop '
'    .HasRetVal = MthLinHasRetVal(MthLin)
    .PmAy = MthLinPmAy(MthLin)
    .RetTy = MthLinRetTy(MthLin)
End With
MthSig = O
End Function
