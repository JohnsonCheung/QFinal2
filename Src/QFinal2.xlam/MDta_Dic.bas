Attribute VB_Name = "MDta_Dic"
Option Explicit
Function DicDrs(A As Dictionary, Optional InclDicValTy As Boolean, Optional Tit$ = "Key Val") As Drs
Dim Fny$()
Fny = SslSy(Tit): If InclDicValTy Then Push Fny, "Val-TypeName"
Set DicDrs = Drs(Fny, DicDry(A, InclDicValTy))
End Function

Function DicDry(A As Dictionary, Optional InclDicValTy As Boolean) As Variant()
Dim I, Dr
If A.Count = 0 Then Exit Function
Dim K(): K = A.Keys
If Sz(K) = 0 Then Exit Function
For Each I In K
    If InclDicValTy Then
        Dr = Array(I, A(I), TypeName(A(I)))
    Else
        Dr = Array(I, A(I))
    End If
    Push DicDry, Dr
Next
End Function

Function DicDry_Dic(DicDry()) As Dictionary
Dim O As New Dictionary
If Sz(DicDry) > 0 Then
   Dim Dr
   For Each Dr In DicDry
       O.Add Dr(0), Dr(1)
   Next
End If
Set DicDry_Dic = O
End Function

Function DicDt(A As Dictionary, Optional DtNm$ = "Dic", Optional InclDicValTy As Boolean) As Dt
Dim Dry()
Dry = DicDry(A, InclDicValTy)
Dim F$
    If InclDicValTy Then
        F = "Key Val Ty"
    Else
        F = "Key Val"
    End If
Set DicDt = Dt(DtNm, F, Dry)
End Function

Function DicFny(Optional InclValTy As Boolean) As String()
DicFny = SslSy("Key Val" & IIf(InclValTy, " Type", ""))
End Function
