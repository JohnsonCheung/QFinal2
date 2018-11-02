Attribute VB_Name = "MIde_Mth_Ix_FT"
Option Explicit

Function SrcMthFTIxAy(A$(), MthNm) As FTIx()
Dim IxAy&(), F&, T&
IxAy = SrcMthNmIxAy(A, MthNm)
Dim J%
For J = 0 To UB(IxAy)
   F = IxAy(J)
   T = SrcMthLx_ToLx(A, F)
   Push SrcMthFTIxAy, FTIx(F, T)
Next
End Function
