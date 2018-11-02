Attribute VB_Name = "MTp_Sq_Pm"
Option Explicit
Private Function FndEr(A() As Lnx, OValidatedPmLy$()) As String()
Dim ErIx1%()
Dim ErIx2%()
    ErIx1 = LnxAy_ErIx_OfDupKey(A)
    ErIx2 = LnxAy_ErIx_OfPfxPercent(A)

Dim ErIx%()
    PushAy ErIx, ErIx1
    PushAy ErIx, ErIx2

Dim ValidatedPmLy$()
    ValidatedPmLy = LnxAy_WhByExclErIxAy(A, ErIx)

End Function

Function FndPm(A() As Lnx, OEr$()) As Dictionary
Dim Ly$()
    OEr = FndEr(A, Ly)
Set FndPm = LinesDicLy_LinesDic(Ly)
End Function
