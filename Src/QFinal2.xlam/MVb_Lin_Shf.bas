Attribute VB_Name = "MVb_Lin_Shf"
Option Explicit

Function ShfBktStr$(OLin$)
Dim O$
O = TakBetBkt(OLin): If O = "" Then Exit Function
ShfBktStr = O
OLin = TakAftBkt(OLin)
End Function

Function ShfChr$(OLin, ChrList$)
Dim F$, P%
F = FstChr(OLin)
P = InStr(ChrList, F)
If P > 0 Then
    ShfChr = Mid(ChrList, P, 1)
    OLin = Mid(OLin, 2)
    Exit Function
End If
End Function

Function ShfLblVal$(OItm$(), QLbl)
'Shf a value in OItm with label being QLbl.
'OItm: is array of LLL or LLL=VVV
'QLbl: is ?LLL or LLL
'If QLbl=?LLL, looking LLL in OItm, return truen and remove it from OItm, else return false only
'If QLbl=LLL,  looking LLL=VVV in OItm, return VVV and remove it from OItm
If FstChr(QLbl) = "?" Then
    Dim Lbl$
    Lbl = RmvFstChr(QLbl)
    If AyHas(OItm, Lbl) Then
        ShfLblVal = 1
        OItm = AyRmvEle(OItm, Lbl)
    Else
        ShfLblVal = 0
    End If
    Exit Function
End If
Dim J%
For J = 0 To UB(OItm)
    With Brk2(OItm(J), "=")
        If .S1 = QLbl Then
            ShfLblVal = .S2
            OItm = AyRmvEleAt(OItm, J)
            Exit Function
        End If
    End With
Next
End Function

Function ShfPfx(OLin, Pfx) As Boolean
If HasPfx(OLin, Pfx) Then
    OLin = RmvPfx(OLin, Pfx)
    ShfPfx = True
End If
End Function

Function ShfPfxSpc(OLin, Pfx) As Boolean
If HasPfxSpc(OLin, Pfx) Then
    OLin = Mid(OLin, Len(Pfx) + 2)
    ShfPfxSpc = True
End If
End Function

Function ShfStarTerm(OItm$(), OLbl$()) As String()
Dim NStar%, I
For Each I In OLbl
    If FstChr(I) <> "*" Then
        If NStar > 0 Then
            OItm = AyMid(OItm, NStar)
            OLbl = AyMid(OLbl, NStar)
            Exit Function
        End If
    End If
    Push ShfStarTerm, OItm(NStar)
    NStar = NStar + 1
Next
End Function

Function ShfVal(OLin$, Lblss$) As String()
'Lin   is: XX YY ZZ=123 [ABC=A 1] ..
'Lblss is: *LL ?LL LL LL
'Return: as many elements as in Lblss
'        *LL means fixed position, no Lbl is required in OLin
'        ?LL means boolean, return 0 or 1, is LL in OLin, return 1, else return 0
'        LL  means match LL=VV in OLin, return VV is match else return ''
'        those element in OLin has matched Lblss will be removed and remaining unmatched will put back into OLin
Dim Lbl$(), Itm$(), Lbli
Lbl = SslSy(Lblss)
Itm = LinTermAy(OLin)
ShfVal = ShfStarTerm(Itm, Lbl)
For Each Lbli In AyNz(Lbl)
    If Sz(Itm) = 0 Then Exit For
    PushI ShfVal, ShfLblVal(Itm, Lbli)
Next
OLin = JnTerm(Itm)
End Function

Sub Z()
Z_ShfBktStr
Z_ShfLblVal
Z_ShfPfx
Z_ShfVal
End Sub

Private Sub Z_ShfBktStr()
Dim A$, Ept1$
A$ = "(O$()) As X": Ept = "O$()": Ept1 = " As X": GoSub Tst
Exit Sub
Tst:
    Act = ShfBktStr(A)
    C
    Ass A = Ept1
    Return
End Sub

Private Sub Z_ShfLblVal()
Dim OItm$(), QLbl, EptOItm$()
OItm = SslSy("A B C=123 D=XYZ")
QLbl = "?B"
Ept = True
EptOItm = SslSy("A C=123 D=XYZ")
GoSub Tst
Exit Sub
Tst:
    Act = ShfLblVal(OItm, QLbl)
    C
    Ass IsEqAy(OItm, EptOItm)
    Return
End Sub

Private Function Z_ShfPfx()
Dim O$: O = "AA{|}BB "
Ass ShfPfx(O, "{|}") = "AA"
Ass O = "BB "
End Function

Private Sub Z_ShfVal()
Dim A$, Lblss$
A = "Txt VTxt=XYZ [Dft=A 1] VRul=123 Req"
Lblss = "*Ty ?Req ?AlwZLen Dft VTxt VRul"
Ept = LinTermAy("Txt 1 1 [A 1] [] 123")
GoSub Tst
Exit Sub
Tst:
    Act = ShfVal(A, Lblss)
    C
    Return
End Sub
