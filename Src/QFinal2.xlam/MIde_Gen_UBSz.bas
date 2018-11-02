Attribute VB_Name = "MIde_Gen_UBSz"
Option Explicit
Private A_TyNm$, A_Md As CodeModule
Sub GenUBSz(MdNm$, TyNm$)
MdGenUBSz Md(MdNm), TyNm
End Sub
Sub MdGenUBSz(A As CodeModule, TyNm$)
Set A_Md = A
A_TyNm = TyNm
Dim ActUB$, EptUB$
Dim ActSz$, EptSz$
Dim UBNm$, SzNm$
    UBNm = TyNm & "UB"
    SzNm = TyNm & "Sz"
    ActUB = MdMthLines(A, UBNm)
    ActSz = MdMthLines(A, SzNm)
    EptUB = ZLinesUB
    EptSz = ZLinesSz
If ActUB <> EptUB Then
    MdMthRmv A, UBNm
    MdLinesApp A, EptUB
End If
If ActSz <> EptSz Then
    MdMthRmv A, SzNm
    MdLinesApp A, EptSz
End If
End Sub

Private Function ZLinesUB$()
Const A$ = "Function ?UB%(A() As ?)" & _
vbCrLf & "?UB = ?Sz(A) - 1" & _
vbCrLf & "End Function"
ZLinesUB = RplQ(A, A_TyNm)
End Function

Private Function ZLinesSz$()
Const A$ = "Function ?Sz%(A() As ?)" & _
vbCrLf & "On Error Resume Next" & _
vbCrLf & "?Sz = UBound(A) + 1" & _
vbCrLf & "End Function"
ZLinesSz = RplQ(A, A_TyNm)
End Function

