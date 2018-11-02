Attribute VB_Name = "MIde_Mth_Ix_TopRmk"
Option Explicit
Function SrcMthIxTopRmk$(A$(), MthIx&)
Dim O$(), J&, L$
Dim Fm&: Fm = SrcMthIxTopRmkFm(A, MthIx)
For J = Fm To MthIx - 1
    L = A(J)
    If FstChr(L) = "'" Then
        If L <> "'" Then
            PushI O, L
        End If
    End If
Next
SrcMthIxTopRmk = Join(O, vbCrLf)
End Function


Function SrcMthIxTopRmkFm&(A$(), MthIx&)
Dim M1&
    Dim J&
    For J = MthIx - 1 To 0 Step -1
        If IsCdLin(A(J)) Then
            M1 = J
            GoTo M1IsFnd
        End If
    Next
    M1 = -1
M1IsFnd:
Dim M2&
    For J = M1 + 1 To MthIx - 1
        If Trim(A(J)) <> "" Then
            M2 = J
            GoTo M2IsFnd
        End If
    Next
    M2 = MthIx
M2IsFnd:
SrcMthIxTopRmkFm = M2
End Function
