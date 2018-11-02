Attribute VB_Name = "MXls_Z_Lo_Fmt_Tit"
Option Explicit

Sub LoTitLySet(A As ListObject, TitLy$())
Dim Sq(), R As Range
    Sq = ZSq(TitLy, LoFny(A)): If Sz(Sq) = 0 Then Exit Sub
    Set R = ZRg(Sq, A)
SqRg Sq, R
ZMge R
RgBdrInner R
RgBdrAround R
End Sub

Private Sub ZMge(A As Range)
Dim J%
For J = 1 To A.Rows.Count
    ZMgeH RgR(A, J)
Next
For J = 1 To A.Columns.Count
    ZMgeV RgC(A, J)
Next
End Sub

Private Sub ZMgeH(A As Range)
A.Application.DisplayAlerts = False
Dim J%, C1%, C2%, V, LasV
LasV = RgRC(A, 1, 1).Value
C1 = 1
For J = 2 To A.Columns.Count
    V = RgRC(A, 1, J).Value
    If V <> LasV Then
        C2 = J - 1
        If Not IsEmpty(LasV) Then
            RgRCC(A, 1, C1, C2).MergeCells = True
        End If
        C1 = J
        LasV = V
    End If
Next
A.Application.DisplayAlerts = True
End Sub

Private Sub ZMgeV(A As Range)
Dim J%
For J = A.Rows.Count To 2 Step -1
    CellMgeAbove RgRC(A, J, 1)
Next
End Sub

Private Function ZRg(Sq(), Lo As ListObject) As Range
Dim At As Range, NR%, R%
R = Lo.DataBodyRange.Row
NR = UBound(Sq, 1)
Set ZRg = RgReSz(RgRC(Lo.DataBodyRange, 0 - NR, 1), Lo)
End Function

Private Function ZSq(TitLy$(), Fny0) As Variant()
Dim Col()
    Dim F, Tit$
    For Each F In SslSy(Fny0)
        Tit = AyFstRmvT1(TitLy, F)
        If Tit = "" Then
            PushI Col, Sy(F)
        Else
            PushI Col, AyTrim(SplitVBar(Tit))
        End If
    Next
ZSq = SqTranspose(DrySq(Col))
End Function

Sub Z()
Z_ZSq
End Sub

Private Sub Z_ZSq()
Dim TitLy$(), Fny0$
'----
Dim A$()
'TitLy
    Erase A
    Push A, "A A1 | A2 11 "
    Push A, "B B1 | B2 | B3"
    Push A, "C C1"
    Push A, "E E1"
    TitLy = A
Fny0 = "A B C D E"
Ept = NewSq(3, 5)
    SqRowDrSet Ept, 1, SslSy("A1 B1 C1 D E1")
    SqRowDrSet Ept, 2, Array("A2 11", "B2")
    SqRowDrSet Ept, 3, Array(Empty, "B3")
GoSub Tst
Exit Sub
'---
'TitLy
    Erase A
    Push A, "ksdf | skdfj  |skldf jf"
    Push A, "skldf|sdkfl|lskdf|slkdfj"
    Push A, "askdfj|sldkf"
    Push A, "fskldf"
SqBrw ZSq(A, "")

Exit Sub
Tst:
    Act = ZSq(TitLy, Fny0)
    Ass SqIsEq(Act, Ept)
    Return
End Sub
