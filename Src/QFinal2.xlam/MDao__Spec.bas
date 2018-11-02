Attribute VB_Name = "MDao__Spec"
Option Explicit
Property Get SpnmFn$(A)
SpnmFn = A & ".txt"
End Property

Property Get SpnmFny() As String()
SpnmFny = DbtFny(CurDb, "Spec")
End Property

Function SpnmFt$(A)
SpnmFt = SpecPth & SpnmFn(A)
End Function

Sub SpnmFtIni(A)
FtIni SpnmFt(A)
End Sub

Sub SpnmImp(A)
DbImpSpec CurDb, A
End Sub

Sub SpnmKill(A)
Kill SpnmFt(A)
End Sub

Property Get SpnmLines$(A)
SpnmLines = SpnmV(A, "Lines")
End Property

Property Let SpnmLines(A, Lines$)
Tfk0V("Spec", "Lines", A) = Lines
End Property

Function SpnmLy(A) As String()
SpnmLy = SplitCrLf(SpnmLines(A))
End Function

Function SpnmTim(A) As Date
SpnmTim = Nz(TfkV("Spec", "Tim", A), 0)
End Function

Property Get SpnmV(A, ValNm$)
SpnmV = TfkV("Spec", ValNm, A)
End Property

Sub SpnmBrw(A)
FtBrw SpnmFt(A)
End Sub

Function SpnmConstLines$(A, Nm$)
'SpnmConstLines = LinesConstLines(SpnmLines(A), Nm)
End Function

Sub SpnmEdt(A)
SpnmEnsRec A
SpnmExpIfFtNotExist A
SpnmBrw A
End Sub

Sub SpnmEnsRec(A)
If TSkIsExist("Spec", A) Then Exit Sub
TkIns "Spec", A
End Sub

Sub SpnmExp(A)
StrWrt SpnmLines(A), SpnmFt(A), OvrWrt:=True
End Sub

Sub SpnmExpIfFtNotExist(A)
If Not FfnIsExist(SpnmFt(A)) Then SpnmExp A
End Sub

Sub SpnmExpIfNotExist(A)
If FfnIsExist(A) Then Exit Sub
StrWrt SpnmLines(A), SpnmFt(A)
End Sub
Sub FmtSpec_Exp()
SpnmExp FmtSpecNm
End Sub

Sub FmtSpec_Imp()
SpnmImp FmtSpecNm
End Sub

Sub FmtSpec_Kill()
SpnmKill FmtSpecNm
End Sub

Function FmtSpec_Lines$()
FmtSpec_Lines = SpnmLines(FmtSpecNm)
End Function

Function FmtSpec_Ly() As String()
FmtSpec_Ly = SplitCrLf(FmtSpec_Lines)
End Function

Function FmtSpecBrw()
SpnmBrw FmtSpecNm
End Function

Function FmtSpecErLy() As String()
Dim A$(), B$(), C$(), D$()
'Dim Fml$(), Wdt$()
'A = LyChk(FmtSpec_Ly, VdtFmtSpecNmSsl)
'B = XFmlChk(Fml)
'C = XWdtChk(Wdt)
'D = XFmtChk(Fmt)
'E = XAlignC(AlignC)
'F = XTSumChk(TSum)
'G = XTAvgChk(TAvg)
'H = XTCntChk(TCnt)
'I = XReSeqChk(ReSeq)
End Function
