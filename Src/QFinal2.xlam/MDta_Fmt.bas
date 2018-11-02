Attribute VB_Name = "MDta_Fmt"
Option Explicit

Private Sub Z_DtFmt()
Dim A As Dt, MaxColWdt%, BrkColNm$, ShwZer As Boolean
'--
Set A = SampleDt1
'Ept = Z_DtFmtEpt1
GoSub Tst
'--
Exit Sub
Tst:
    Act = DtFmt(A, MaxColWdt, BrkColNm, ShwZer)
    C
    Return
End Sub

Function DtFmt(A As Dt, Optional MaxColWdt% = 100, Optional BrkColNm$, Optional ShwZer As Boolean, Optional HidIxCol As Boolean) As String()
Push DtFmt, "*Tbl " & A.DtNm
PushAy DtFmt, DrsFmt(DtDrs(A), MaxColWdt, BrkColNm, ShwZer, HidIxCol)
End Function

Function DsFmt(A As Ds, Optional MaxColWdt% = 100, Optional DtBrkLinMapStr$, Optional HidIxCol As Boolean) As String()
Push DsFmt, "*Ds " & A.DsNm & "=================================================="
Dim Dic As Dictionary ' DicOf_Tn_to_BrkColNm
    Set Dic = LyDic(DtBrkLinMapStr)
Dim J%, DtNm$, Dt As Dt, BrkColNm$, DtAy() As Dt
DtAy = A.DtAy
For J = 0 To UB(DtAy)
    Set Dt = DtAy(J)
    DtNm$ = Dt.DtNm
    If Dic.Exists(DtNm) Then BrkColNm = Dic(DtNm) Else BrkColNm = ""
    PushAy DsFmt, DtFmt(Dt, MaxColWdt, BrkColNm, HidIxCol)
Next
End Function

Function DrsFmt(A As Drs, Optional MaxColWdt% = 100, Optional BrkColNm$, Optional ShwZer As Boolean, Optional HidIxCol As Boolean) As String()
'If BrkColNm changed, insert a break line if BrkColNm is given
Dim Drs As Drs
    If HidIxCol Then
        Set Drs = A
    Else
        Set Drs = DrsAddRowIxCol(A)
    End If
Dim BrkColIx%
    BrkColIx = AyIx(A.Fny, BrkColNm)
    If BrkColIx >= 0 Then
        If Not HidIxCol Then
            BrkColIx = BrkColIx + 1 ' Need to increase by 1 due the Ix column is added
        End If
    End If
Dim Dry()
    Dry = Drs.Dry
    PushI Dry, Drs.Fny
Dim Ay$()
    Ay = DryFmt(Dry, MaxColWdt, BrkColIx, ShwZer) '<== Will insert break line if BrkColIx>=0

Dim U&: U = UB(Ay)
Dim Hdr$: Hdr = Ay(U - 1)
Dim Lin$: Lin = Ay(U)
DrsFmt = AyRmvLasNEle(AyInsAy(Ay, Array(Lin, Hdr, Lin)), 2)
PushI DrsFmt, Lin
End Function

Sub Z_DsFmt()
D DsFmt(SampleDs)
End Sub
