Attribute VB_Name = "MIde_Mth_Ix"
Option Explicit
Sub Z()
Z_SrcMthIx
Z_SrcMthIx1
End Sub
Private Sub Z_SrcMthIx()
Dim IxAy&(), Src$(), Ix
Src = MdSrc(Md("AAAMod"))
IxAy = SrcMthIx(Src)
For Each Ix In IxAy
    If LinMthKd(Src(Ix)) = "" Then
        Debug.Print Ix
        Debug.Print Src(Ix)
    End If
Next
End Sub

Function MthFC(A As Mth) As FmCnt()
MthFC = SrcMthNmFC(MdBdyLy(A.Md), A.Nm)
End Function


Function SrcMthIx(A$(), Optional B As WhMth) As Long()
Dim J&
For J = 0 To UB(A)
    If LinIsMthWh(A(J), B) Then
        PushI SrcMthIx, J
    End If
Next
End Function

Function SrcMthNmIxAy(A$(), MthNm) As Long()
Dim L, J&, Ix&
Ix = SrcMthNmIx(A, MthNm): If Ix = -1 Then Exit Function
PushI SrcMthNmIxAy, Ix
If LinIsPrp(A(Ix)) Then
    Ix = SrcMthNmIx(A, MthNm, Ix + 1)
    If Ix > 0 Then
        PushI SrcMthNmIxAy, Ix
    End If
End If
End Function
Function SrcMthNmIx&(A$(), MthNm, Optional Fm& = 0)
Dim I
For I = Fm To UB(A)
    If LinMthNm(A(I)) = MthNm Then
        SrcMthNmIx = I
        Exit Function
    End If
Next
SrcMthNmIx = -1
End Function
Function SrcMthIxTo&(A$(), MthIx&)
Dim T$, F$, J&
T = LinMthKd(A(MthIx)): If T = "" Then Stop
F = "End " & T
For J = MthIx + 1 To UB(A)
    If HasPfx(A(J), F) Then SrcMthIxTo = J: Exit Function
Next
Stop
End Function

Private Sub Z_SrcMthIx1()
Dim A$(), Ix&(), O$(), I
A = CurSrc
Ix = SrcMthIx(CurSrc)
For Each I In Ix
    PushI O, A(I)
Next
Brw O
End Sub

Function SrcFstMthIx&(A$())
Dim J&
For J = 0 To UB(A)
   If LinIsMth(A(J)) Then
       SrcFstMthIx = J
       Exit Function
   End If
Next
SrcFstMthIx = -1
End Function
