Attribute VB_Name = "MVb_Dic_Cmp"
Option Explicit
Private Type CmpRslt
    AExcess As Dictionary
    BExcess As Dictionary
    ADif As Dictionary
    BDif As Dictionary
    Sam As Dictionary
End Type
Private A As Dictionary, B As Dictionary
Private A_Nm$, B_Nm$
Private Sub Z_DicCmpBrw()
Set A = DicVblDic("X AA|A BBB|A Lines1|A Line3|B Line1|B line2|B line3..")
Set B = DicVblDic("X AA|C Line|D Line1|D line2|B Line1|B line2|B line3|B Line4")
DicCmpBrw A, B
End Sub
Private Function ZCmp(A_Dic As Dictionary, B_Dic As Dictionary, ANm$, BNm$) As CmpRslt
Set A = A_Dic
Set B = B_Dic
A_Nm = ANm
B_Nm = BNm
With ZCmp
    Set .AExcess = DicMinus(A, B)
    Set .BExcess = DicMinus(B, A)
    Set .Sam = ZC_Sam
        ZC_SamKey .ADif, .BDif
End With
End Function
Function DicCmpFmt(A As Dictionary, B As Dictionary, Optional Nm1$ = "Fst", Optional Nm2$ = "Snd") As String()
DicCmpFmt = ZFmt(ZCmp(A, B, Nm1, Nm2))
End Function
Sub DicCmpBrw(A As Dictionary, B As Dictionary, Optional Nm1$ = "Fst", Optional Nm2$ = "Snd")
AyBrw DicCmpFmt(A, B, Nm1, Nm2)
End Sub

Private Function ZFmt(A As CmpRslt) As String()
With A
    ZFmt = AyAddAp( _
        ZF_Excess(.AExcess, A_Nm), _
        ZF_Excess(.BExcess, B_Nm), _
        ZF_Dif(.ADif, .BDif), _
        ZF_Sam(.Sam))
End With
End Function

Private Function ZF_Excess(A As Dictionary, Nm$) As String()
If A.Count = 0 Then Exit Function
Dim K, S1$, S2$, S(0) As S1S2
S2 = "!" & "Er Excess (" & Nm & ")"
For Each K In A.Keys
    S1 = K & vbCrLf & LinesUnderLin(K) & vbCrLf & A(K)
    Set S(0) = S1S2(S1, S2)
    PushAy ZF_Excess, S1S2AyFmt(S)
Next
End Function

Private Function ZF_Dif(A As Dictionary, B As Dictionary) As String()
If A.Count <> B.Count Then Stop
If A.Count = 0 Then Exit Function
Dim O$(), K, S1$, S2$, S(0) As S1S2, Ly$(), KK$
For Each K In A
    KK = K
    S1 = "!" & "Er Dif" & vbCrLf & K & vbCrLf & LinesUnderLin(KK) & vbCrLf & A(K)
    S2 = "!" & "Er Dif" & vbCrLf & K & vbCrLf & LinesUnderLin(KK) & vbCrLf & B(K)
    Set S(0) = S1S2(S1, S2)
    Ly = S1S2AyFmt(S)
    PushAy O, Ly
Next
ZF_Dif = O
End Function

Private Function ZF_Sam(A As Dictionary) As String()
If A.Count = 0 Then Exit Function
Dim O$(), K, S() As S1S2, KK$
For Each K In A.Keys
    KK = K
    PushObj S, S1S2("*Same", K & vbCrLf & LinesUnderLin(KK) & vbCrLf & A(K))
Next
ZF_Sam = S1S2AyFmt(S)
End Function

Private Sub ZC_SamKey( _
    OADif As Dictionary, OBDif As Dictionary)
Dim K
Set OADif = New Dictionary
Set OBDif = New Dictionary
For Each K In A.Keys
    If B.Exists(K) Then
        If A(K) <> B(K) Then
            OADif.Add K, A(K)
            OBDif.Add K, B(K)
        End If
    End If
Next
End Sub

Private Function ZC_Sam() As Dictionary
Set ZC_Sam = New Dictionary
If A.Count = 0 Or B.Count = 0 Then Exit Function
Dim K
For Each K In A.Keys
    If B.Exists(K) Then
        If A(K) = B(K) Then
            ZC_Sam.Add K, A(K)
        End If
    End If
Next
End Function

