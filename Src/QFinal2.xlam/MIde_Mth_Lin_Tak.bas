Attribute VB_Name = "MIde_Mth_Lin_Tak"
Option Explicit

Function TakMdy$(A)
TakMdy = TakPfxAySpc(A, MdyAy)
End Function

Function TakMthKd$(A)
TakMthKd = TakPfxAySpc(A, MthKdAy)
End Function

Function TakMthShtTy$(A)
Dim B$
B = TakPfxAy(A, MthTyAy): If B = "" Then Exit Function
TakMthShtTy = MthShtTy(B)
End Function

Function TakMthTy$(A)
TakMthTy = TakPfxAySpc(A, MthTyAy)
End Function

