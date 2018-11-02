Attribute VB_Name = "MDao_CurDb_TT"
Option Explicit
Sub TTAddPfx(TT$, Pfx$)
DbttAddPfx CurDb, TT, Pfx
End Sub

Sub TTBrw(TT$)
DbttBrw CurDb, TT
End Sub

Sub TTDrp(TT$)
DbttDrp CurDb, TT
End Sub

Sub TTOpn(TT)
AyDo SslSy(TT), "TblOpn"
End Sub

Function TTScly(Tny0) As String()
TTScly = AySy(AyOfAy_Ay(AyMap(CvNy(Tny0), "TblScly")))
End Function
