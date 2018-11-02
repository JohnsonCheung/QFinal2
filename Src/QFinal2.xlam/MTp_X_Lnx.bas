Attribute VB_Name = "MTp_X_Lnx"
Option Explicit
Sub LnxAyBrw(A() As Lnx)
AyBrw LnxAyFmt(A)
End Sub
Function LnxAyFmt(A() As Lnx) As String()
Dim I
For Each I In AyNz(A)
    With CvLnx(I)
        PushI LnxAyFmt, "L#(" & .Ix & ") " & .Lin
    End With
Next
End Function
Function LnxAyLy(A() As Lnx) As String()
LnxAyLy = OyPrpSy(A, "Lin")
End Function
Function LnxRmvT1$(A As Lnx)
If Not IsNothing(A) Then LnxRmvT1 = RmvT1(A.Lin)
End Function
Sub LnxAsg(A As Lnx, OLin$, OIx%)
With A
    OLin = .Lin
    OIx = .Ix
End With
End Sub
Function LnxAyT1Chk(A() As Lnx, T1Ay0) As String()
Dim T1Ay$(), T1$, L, O$()
T1Ay = CvNy(T1Ay0)
For Each L In A
    T1 = LinT1(L)
    If Not AyHas(T1Ay, LinT1(L)) Then Push O, L
Next
If Sz(O) > 0 Then
    O = AyAddPfx(AyQuoteSqBkt(O), Space(4))
    O = AyIns(O, FmtQQ("Following lines have invalid T1.  Valid T1 are [?]", JnSpc(T1Ay)))
End If
LnxAyT1Chk = O
End Function

Function LnxAyWhRmvT1(A() As Lnx, T1) As Lnx()
Dim O()  As Lnx, X
For Each X In AyNz(A)
    With CvLnx(X)
        If LinT1(.Lin) = T1 Then
            PushObj O, Lnx(.Ix, RmvT1(.Lin))
        End If
    End With
Next
LnxAyWhRmvT1 = O
End Function

Function LnxStr$(A As Lnx)
LnxStr = "L#" & A.Ix + 1 & ": " & A.Lin
End Function
Function Lnx(Ix, Lin) As Lnx
Set Lnx = New Lnx
With Lnx
    .Lin = Lin
    .Ix = Ix
End With
End Function


Function CvLnx(A) As Lnx
Set CvLnx = A
End Function
