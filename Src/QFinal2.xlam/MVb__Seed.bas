Attribute VB_Name = "MVb__Seed"
Option Explicit
Function Seed_Expand$(VblQQStr, Ny0)
'Seed is a VblQQ-String
Dim A$, J%, O$()
Dim Ny$()
Ny = CvNy(Ny0)
For J = 0 To UB(Ny)
    Push O, Replace(VblQQStr, "?", Ny(J))
Next
Seed_Expand = RplVBar(JnCrLf(O))
End Function

Function SeedExpand$(QVbl$, Ny$())
Dim O$()
Dim Sy$(): Sy = SplitVBar(QVbl)
Dim J%, I
For J = 0 To UB(Ny)
    For Each I In Sy
       Push O, Replace(I, "?", Ny(J))
    Next
Next
SeedExpand = JnCrLf(O)
End Function

Private Sub Z_SeedExpand()
Dim Tp$
Dim Seed$()
Tp = "Sub Tst?()|Dim A As New ?: A.Tst|End Sub"
Seed = SslSy("Xws Xwb Xfx Xrg")
'Debug.Print SeedExpand(Seed, Tp)
End Sub

Private Sub ZZ_Seed_Expand()
Dim Ny0
Dim QVbl$
QVbl = "Sub Tst?()|Dim A As New ?: A.Tst|End Sub"
Ny0 = "Xws Xwb Xfx Xrg"
Debug.Print Seed_Expand(QVbl, Ny0)
End Sub
