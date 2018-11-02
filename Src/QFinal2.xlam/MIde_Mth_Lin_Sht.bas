Attribute VB_Name = "MIde_Mth_Lin_Sht"
Option Explicit

Function ShtMdy$(Mdy)
Select Case Mdy
Case "": Exit Function
Case "Private": ShtMdy = "Prv"
Case "Friend": ShtMdy = "Frd"
Case "Public": ShtMdy = "Pub"
End Select
End Function

Function ShtMdyAy() As String()
Static O$(2), A As Boolean
If Not A Then
    A = True
    O(0) = "Prv"
    O(1) = "Frd"
    O(2) = "Pub"
End If
ShtMdyAy = O
End Function
