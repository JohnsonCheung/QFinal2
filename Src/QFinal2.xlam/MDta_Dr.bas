Attribute VB_Name = "MDta_Dr"
Option Explicit


Sub DrPutSq(A, Sq, Optional R& = 1, Optional NoTxtSngQ As Boolean)
Dim J%, I
If NoTxtSngQ Then
    For Each I In A
        J = J + 1
        Sq(R, J) = I
    Next
    Exit Sub
End If
For Each I In A
    J = J + 1
    If IsStr(I) Then
        Sq(R, J) = "'" & I
    Else
        Sq(R, J) = I
    End If
Next
End Sub

Function DrFmt$(Dr, Wdt%(), Optional Sep$ = " | ")
Dim UDr%
   UDr = UB(Dr)
Dim O$()
   Dim U1%: U1 = UB(Wdt)
   ReDim O(U1)
   Dim W, V
   Dim J%, V1$
   J = 0
   For Each W In Wdt
       If UDr >= J Then V = Dr(J) Else V = ""
       V1 = AlignL(V, W)
       O(J) = V1
       J = J + 1
   Next
If Sep = " | " Then
    DrFmt = "| " & Join(O, Sep) & " |"
Else
    DrFmt = Join(O, Sep)
End If
End Function
