Attribute VB_Name = "MVb_Run_FunTim"
Option Explicit
Sub FunTim(FunNy0)
Dim B!, E!, F
For Each F In CvNy(FunNy0)
    B = Timer
    Run F
    E = Timer
    Debug.Print F, "<-- Run"; E - B
Next
End Sub

Private Sub ZZ_FunTim()
FunTim "ZZA ZZB"
End Sub

Private Sub ZZA()
Dim J&, I&
For J = 0 To 100
    For I = 0 To 100
        Debug.Print I
    Next
Next
End Sub
Private Sub ZZB()
Dim J&, I&
For J = 0 To 100
    For I = 0 To 100
        Debug.Print I
    Next
Next
End Sub
