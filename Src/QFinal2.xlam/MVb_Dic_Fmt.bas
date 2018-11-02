Attribute VB_Name = "MVb_Dic_Fmt"
Option Explicit
Sub DicDmp(A As Dictionary, Optional InclDicValTy As Boolean, Optional Tit$ = "Key Val")
D DicFmt(A, InclDicValTy, Tit)
End Sub
Function DicFmt(A As Dictionary, Optional InclValTy As Boolean, Optional Tit$ = "Key Val") As String()
If ZHasLines(A) Then
    DicFmt = S1S2AyFmt(DicS1S2Ay(A))
Else
    DicFmt = ZLinFmt(A, InclValTy)
End If
End Function
Private Function ZHasLines(A As Dictionary) As Boolean
ZHasLines = True
Dim K
For Each K In A.Keys
    If IsLines(K) Then Exit Function
    If IsLines(A(K)) Then Exit Function
Next
ZHasLines = False
End Function
Private Function ZLinFmt(A As Dictionary, Optional InclDicValTy As Boolean) As String()
Dim K, O$()
If InclDicValTy Then ZLinFmt = ZLinFmt1(A): Exit Function
For Each K In A.Keys
    PushI O, K & " " & A(K)
Next
ZLinFmt = AyAlign1T(O)
End Function
Private Function ZLinFmt1(A As Dictionary) As String()
Dim K, O$()
For Each K In A.Keys
    PushI O, K & " " & TypeName(A(K)) & " " & A(K)
Next
ZLinFmt1 = AyAlign2T(O)
End Function
Sub DicBrw(A As Dictionary, Optional InclDicValTy As Boolean)
AyBrw DicFmt(A, InclDicValTy)
End Sub
