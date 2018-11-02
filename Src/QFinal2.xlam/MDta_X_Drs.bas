Attribute VB_Name = "MDta_X_Drs"
Option Explicit
Function Z_DrsFmt()
GoTo ZZ
ZZ:
AyDmp DrsFmt(SampleDrs1)
End Function
Sub PushDrs(O As Drs, A As Drs)
If IsNothing(O) Then
    Set O = A
    Exit Sub
End If
If IsNothing(A) Then Exit Sub
If Not IsEq(O.Fny, A.Fny) Then Stop
Set O = Drs(O.Fny, CvAy(AyAddAp(O.Dry, A.Dry)))
End Sub

Private Sub ZZ_DrsFmt_InsBrkLin()
Dim TblLy$()
Dim Act$()
Dim Exp$()
'TblLy = FtLy(TstResPth & "DrsFmt_InsBrkLin.txt")
'Act = DrsFmt_InsBrkLin(TblLy, "Tbl")
'Exp = FtLy(TstResPth & "DrsFmt_InsBrkLin_Exp.txt")
'AyPair_EqChk Exp, Act
End Sub

Private Sub ZZ_DrsGpDic()
Dim Act As Dictionary, Dry(), Dr1, Dr2, Dr3
Dr1 = Array("A", , 1)
Dr2 = Array("A", , 2)
Dr3 = Array("B", , 3)
Dry = Array(Dr1, Dr2, Dr3)
Set Act = DryGpDic(Dry, 0, 2)
Ass Act.Count = 2
Ass IsEqAy(Act("A"), Array(1, 2))
Ass IsEqAy(Act("B"), Array(3))
Stop
End Sub

Function CvDrs(A) As Drs
Set CvDrs = A
End Function

Function Drs(Fny0, Dry()) As Drs
Dim O As New Drs
Set Drs = O.Init(CvNy(Fny0), Dry)
End Function

Private Sub ZZ_DrsGpFlat()
Dim Act As Drs, Drs2 As Drs, Drs1 As Drs, N1%, N2%
'Set Drs1 = VbeFun12Drs(CurVbe)
N1 = Sz(Drs1.Dry)
'Set Drs2 = VbeMth12Drs(CurVbe)
'N2 = Sz(Drs2.Dry)
'Debug.Print N1, N2
Set Act = DrsGpFlat(Drs1, "Nm", "Lines")
DrsBrw Act
End Sub

Private Sub ZZ_DrsGpFlat_1()
Dim Act As Drs, D As Drs, Dr1, Dr2, Dr3
Dr1 = Array("A", , 1)
Dr2 = Array("A", , 2)
Dr3 = Array("B", , 3)
Set D = Drs("A B C", CvAy(Array(Dr1, Dr2, Dr3)))
Set Act = DrsGpFlat(D, "A", "C")
Stop
DrsBrw Act
End Sub

Private Sub ZZ_DrsKeyCntDic()
Dim Drs As Drs, Dic As Dictionary
'Set Drs = VbeMth12Drs(CurVbe)
Set Dic = DrsKeyCntDic(Drs, "Nm")
DicBrw Dic
End Sub

Private Sub ZZ_DrsSel()
DrsBrw DrsSel(SampleDrs1, "A B D")
End Sub

Function DrsAddCol(A As Drs, ColNm$, ColVal) As Drs
Dim Fny$(): Fny = A.Fny
Dim NewFny$(): NewFny = Fny: PushI NewFny, ColNm
Set DrsAddCol = Drs(NewFny, DryAddCol(A.Dry, ColVal))
End Function

Function DrsAddConstCol(A As Drs, ColNm$, ConstVal) As Drs
Dim Fny$()
    Fny = A.Fny
    Push Fny, ColNm
Set DrsAddConstCol = Drs(Fny, DryAddConstCol(A.Dry, ConstVal))
End Function

Function DrsIsEq(A As Drs, B As Drs) As Boolean
If Not IsEqAy(A.Fny, B.Fny) Then Exit Function
If Not DryIsEq(A.Dry, B.Dry) Then Exit Function
DrsIsEq = True
End Function

Function DrsAddRowIxCol(A As Drs) As Drs
Dim Fny$()
Dim Dry()
    Fny = AyIns(A.Fny, "Ix")
    Dim J&, Dr
    For Each Dr In AyNz(A.Dry)
        Dr = AyIns(Dr, J): J = J + 1
        Push Dry, Dr
    Next
Set DrsAddRowIxCol = Drs(Fny, Dry)
End Function

Function DrsAddValIdCol(A As Drs, ColNm$, Optional ColNmPfx$) As Drs
Dim Ix%, Fny$()
Fny = A.Fny
Ix = AyIx(Fny, ColNm): If Ix = -1 Then Stop
    Dim X$, Y$, C$
        C = ColNmPfx & ColNm
        X = C & "Id"
        Y = C & "Cnt"
    If AyHas(Fny, X) Then Stop
    If AyHas(Fny, Y) Then Stop
    PushIAy Fny, Array(X, Y)
Set DrsAddValIdCol = Drs(Fny, DryAddValIdCol(A.Dry, Ix))
End Function

Sub DrsBrw(A As Drs, Optional MaxColWdt% = 100, Optional BrkColNm$, Optional Fnn$)
AyBrw DrsFmt(A, MaxColWdt, BrkColNm$), Fnn
End Sub

Function DrsCol(A As Drs, ColNm) As Variant()
'DrsCol = DryColInto(A.Dry, ColNm)
End Function

Function DrsColInto(A As Drs, F, OInto)
Dim O, Ix%, Dry(), Dr
Ix = AyIx(A.Fny, F): If Ix = -1 Then Stop
O = OInto
Erase O
Dry = A.Dry
If Sz(Dry) = 0 Then DrsColInto = O: Exit Function
For Each Dr In Dry
    Push O, Dr(Ix)
Next
DrsColInto = O
End Function

Function DrsColSy(A As Drs, F) As String()
DrsColSy = DrsColInto(A, F, EmpSy)
End Function

Sub DrsDmp(A As Drs, Optional MaxColWdt% = 100, Optional BrkColNm$)
AyDmp DrsFmt(A, MaxColWdt, BrkColNm$)
End Sub

Function DrsDrpCol(A As Drs, ColNy0) As Drs
Dim ColNy$(): ColNy = CvNy(ColNy0)
Ass AyHasSubAy(A.Fny, ColNy)
Dim IxAy&()
    IxAy = AyIxAy(A.Fny, ColNy)
Dim Fny$(), Dry()
    Fny = AyWhExlIxAy(A.Fny, IxAy)
    Dry = DryRmvColByIxAy(A.Dry, IxAy)
Set DrsDrpCol = Drs(Fny, Dry)
End Function


Function DrsFF$(A As Drs)
DrsFF = JnSpc(A.Fny)
End Function

Function LblSeqAy(A, N%) As String()
Dim O$(), J%, F$, L%
L = Len(N)
F = StrDup("0", L)
ReDim O(N - 1)
For J = 1 To N
    O(J - 1) = A & Format(J, F)
Next
LblSeqAy = O
End Function

Function LblSeqSsl$(A, N%)
LblSeqSsl = Join(LblSeqAy(A, N), " ")
End Function

Function DrsGpFlat(A As Drs, K$, G$) As Drs
Dim Fny0$, Dry(), S$, N%, Ix%()
Ix = AyIxAyI(A.Fny, Array(K, G))
Dry = DryGpFlat(A.Dry, Ix(0), Ix(1))
N = DryNCol(Dry) - 2
S = LblSeqSsl(G, N)
Fny0 = FmtQQ("? N ?", K, S)
Set DrsGpFlat = Drs(Fny0, Dry)
End Function

Function DrsInsCol(A As Drs, ColNm$, C) As Drs
Set DrsInsCol = Drs(AyIns(A.Fny, ColNm), DryInsCol(A.Dry, C))
End Function

Function DrsInsColAft(A As Drs, C$, FldNm$) As Drs
Set DrsInsColAft = DrsInsColXxx(A, C, FldNm, True)
End Function

Function DrsInsColBef(A As Drs, C$, FldNm$) As Drs
Set DrsInsColBef = DrsInsColXxx(A, C, FldNm, False)
End Function

Private Function DrsInsColXxx(A As Drs, C$, FldNm$, IsAft As Boolean) As Drs
Dim Fny$(), Dry(), Ix&, Fny1$()
Fny = A.Fny
Ix = AyIx(Fny, C): If Ix = -1 Then Stop
If IsAft Then
    Ix = Ix + 1
End If
Fny1 = AyIns(Fny, FldNm, CLng(Ix))
Dry = DryInsCol(A.Dry, Ix)
Set DrsInsColXxx = Drs(Fny1, Dry)
End Function

Function DrsKeyCntDic(A As Drs, K$) As Dictionary
Dim Dry(), O As New Dictionary, Fny$(), Dr, Ix%, KK$
Fny = A.Fny
Ix = AyIx(Fny, K)
Dry = A.Dry
If Sz(Dry) > 0 Then
    For Each Dr In A.Dry
        KK = Dr(Ix)
        If O.Exists(KK) Then
            O(KK) = O(KK) + 1
        Else
            O.Add KK, 1
        End If
    Next
End If
Set DrsKeyCntDic = O
End Function

Function DrsLines_Drs(A) As Drs
'Set DrsLines_Drs = DrsLy_Drs(SplitLines(A))
End Function

Function DrsNCol%(A As Drs)
DrsNCol = Max(Sz(A.Fny), DryNCol(A.Dry))
End Function

Function DrsNRow&(A As Drs)
DrsNRow = Sz(A.Dry)
End Function

Function DrsPkDiff(A As Drs, B As Drs, PkSs$) As Drs

End Function

Function DrsPkMinus(A As Drs, B As Drs, PkSs$) As Drs
Dim Fny$(), PkIxAy&()
Fny = A.Fny: If Not IsEqAy(Fny, B.Fny) Then Stop
PkIxAy = AyIxAy(Fny, SslSy(PkSs))
Set DrsPkMinus = Drs(Fny, DryPkMinus(A.Dry, B.Dry, PkIxAy))
End Function

Function DrsReOrd(A As Drs, Partial_Fny0) As Drs
Dim ReOrdFny$(): ReOrdFny = CvNy(Partial_Fny0)
Dim IxAy&(): IxAy = AyIxAy(A.Fny, ReOrdFny)
Dim OFny$(): OFny = AyReOrd(A.Fny, IxAy)
Dim ODry(): ODry = DryReOrd(A.Dry, IxAy)
Set DrsReOrd = Drs(OFny, ODry)
End Function

Function DrsRowCnt&(A As Drs, ColNm$, EqVal)
DrsRowCnt = DryRowCnt(A.Dry, AyIx(A.Fny, ColNm), EqVal)
End Function


Function DrsSq(A As Drs) As Variant()
Dim NC&, NR&, Dry(), Fny$()
    Fny = A.Fny
    Dry = A.Dry
    NC = Max(DryNCol(Dry), Sz(Fny))
    NR = Sz(Dry)
Dim O()
ReDim O(1 To 1 + NR, 1 To NC)
Dim C&, R&, Dr()
    For C = 1 To Sz(Fny)
        O(1, C) = Fny(C - 1)
    Next
    For R = 1 To NR
        Dr = Dry(R - 1)
        For C = 1 To Min(Sz(Dr), NC)
            O(R + 1, C) = Dr(C - 1)
        Next
    Next
DrsSq = O
End Function

Function DrsSy(A As Drs, ColNm) As String()
DrsSy = DrsStrCol(A, ColNm)
End Function

Function DrsStrCol(A As Drs, ColNm) As String()
DrsStrCol = AySy(DrsCol(A, ColNm))
End Function

Function DrsVbl_Drs(DrsVbl$) As Drs
'SpecStr:Vbl:VbarLine
'SpecStr:DrsVbl:Data-record-set-vbar-line
'DrsVbl_Drs = DrsLy_Drs(SplitVBar(DrsVbl))
Stop '
End Function

Function DrsDt(A As Drs, DtNm$) As Dt
Set DrsDt = Dt(DtNm, A.Fny, A.Dry)
End Function

Function DrsLinesDrs(DrsLines$) As Drs
Set DrsLinesDrs = DrsLyDrs(SplitCrLf(DrsLines))
End Function

Function DrsLyDrs(DrsLy$()) As Drs
Dim J&, Dry()
If Sz(DrsLy) = 0 Then Exit Function
For J = 1 To UB(DrsLy)
    PushI Dry, LinTermAy(DrsLy(J))
Next
Set DrsLyDrs = Drs(LinTermAy(DrsLy(0)), Dry)
End Function

Function DrsVblDrs(DrsVbl$) As Drs
Set DrsVblDrs = DrsLyDrs(SplitVBar(DrsVbl))
End Function
