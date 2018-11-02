Attribute VB_Name = "MDta_Wh"
Option Explicit

Function DrsWh(A As Drs, Fld, V) As Drs
Set DrsWh = Drs(A.Fny, DryWh(A.Dry, AyIx(A.Fny, Fld), V))
End Function

Function DrsWhCCNe(A As Drs, C1$, C2$) As Drs
Dim Fny$()
Fny = A.Fny
Set DrsWhCCNe = Drs(Fny, DryWhCCNe(A.Dry, AyIx(Fny, C1), AyIx(Fny, C2)))
End Function

Function DrsWhColEq(A As Drs, C$, V) As Drs
Dim Dry(), Ix%, Fny$()
Fny = A.Fny
Ix = AyIx(Fny, C)
Set DrsWhColEq = Drs(Fny, DryWhColEq(A.Dry, Ix, V))
End Function

Function DrsWhColGt(A As Drs, C$, V) As Drs
Dim Dry(), Ix%, Fny$()
Fny = A.Fny
Ix = AyIx(Fny, C)
Set DrsWhColGt = Drs(Fny, DryWhColGt(A.Dry, Ix, V))
End Function

Function DrsWhNotIx(A As Drs, IxAy&()) As Drs
Dim ODry(), Dry()
    Dry = A.Dry
    Dim J&, I&
    For J = 0 To UB(A.Dry)
        If Not AyHas(IxAy, J) Then
            PushI ODry, Dry(J)
        End If
    Next
DrsWhNotIx = Drs(A.Fny, ODry)
End Function

Function DrsWhNotRowIxAy(A As Drs, RowIxAy&()) As Drs
Dim O(), Dry()
    Dry = A.Dry
    Dim J&
    For J = 0 To UB(Dry)
        If Not AyHas(RowIxAy, J) Then
            Push O, Dry(J)
        End If
    Next
Set DrsWhNotRowIxAy = Drs(A.Fny, O)
End Function

Function DrsWhRowIxAy(A As Drs, RowIxAy&()) As Drs
Dim O()
    Dim I, Dry()
    Dry = A.Dry
    For Each I In AyNz(RowIxAy)
        Push O, Dry(I)
    Next
Set DrsWhRowIxAy = Drs(A.Fny, O)
End Function
