Attribute VB_Name = "MVb_Ay_Dry"
Option Explicit

Function AyXCDry(A, C) As Variant()
'XCDry is AyItmX-Const-Dry
Dim X
For Each X In AyNz(A)
    PushI AyXCDry, Array(X, C)
Next
End Function

Function AyCXDry(A, C) As Variant()
'CXDry is Const-AyItmX-Dry
Dim X
For Each X In AyNz(A)
    PushI AyCXDry, Array(C, X)
Next
End Function
