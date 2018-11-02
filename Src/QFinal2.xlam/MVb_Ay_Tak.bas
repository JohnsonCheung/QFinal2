Attribute VB_Name = "MVb_Ay_Tak"
Option Explicit

Function AyTakBefDD(A) As String()
AyTakBefDD = AyMapSy(A, "TakBefDD")
End Function

Function AyTakBefDot(A) As String()
Dim X
For Each X In AyNz(A)
    PushI AyTakBefDot, TakBefDot(X)
Next
End Function

Function AyTakBefOrAll(A, Sep$) As String()
Dim I
For Each I In AyNz(A)
    Push AyTakBefOrAll, TakBefOrAll(I, Sep)
Next
End Function

Function AyTakT1(A) As String()
Dim L
For Each L In AyNz(A)
    PushI AyTakT1, LinT1(L)
Next
End Function

Function AyTakT2(A) As String()
AyTakT2 = AyMapSy(A, "LinT2")
End Function

Function AyTakT3(A) As String()
AyTakT3 = AyMapSy(A, "LinT3")
End Function
