Attribute VB_Name = "MIde__Fxa"
Option Explicit
Option Compare Text
Const CMod$ = "LibIdeFxa."

Sub FxaCrt(A)
WbSavAs(CurXls.Workbooks.Add, A, XlFileFormat.xlOpenXMLAddIn).Close False
End Sub

Function FxaNmFxa$(A)
If FstChr(A) <> "Q" Then Stop
FxaNmFxa = CurPjPth & A & ".xlam"
End Function

Function FxaNmPj(A) As VBProject
Set FxaNmPj = FxaPj(FxaNmFxa(A))
End Function

Function FxaOpn(A) As VBProject
Dim V As Vbe
Dim PjNm$
    Set V = CurXls.Vbe
    PjNm = FxaPjNm(A)
If Not VbeHasPj(V, PjNm) Then
    CurXls.Workbooks.Open A
End If
Set FxaOpn = VbePj(V, PjNm)
End Function

Function FxaPj(A) As VBProject
Const CSub$ = CMod & "FxaCrt"
AssNoFfn A, "Fxa", CSub
If Not IsFxa(A) Then Er CSub, "Given [Fxa] is not ends with .xlam and file name not begins with Q", A
FxaCrt A
Set FxaPj = FxaSetPjNm(A)
End Function

Function FxaPjNm$(A)
FxaPjNm = FfnFnn(A)
End Function

Function FxaSetPjNm(A) As VBProject
CurXls.Workbooks.Open A
Dim O As VBProject
Set O = VbePjFfnPj(CurXls.Vbe, A)
O.Name = FfnFnn(A)
PjSav O
Set FxaSetPjNm = O
End Function

Function IsFxa(A) As Boolean
IsFxa = LCase(FfnExt(A)) = ".xlam"
End Function

Function PjIsFxa(A As VBProject) As Boolean
PjIsFxa = IsFxa(PjFfn(A))
End Function

Sub SrcPthBldFxa(SrcPth$)
Dim P As VBProject
   Dim Fnn$, F$
   Fnn = FfnFnn(RmvLasChr(SrcPth))
   F = SrcPth & Fnn & ".xlam"
   Set P = FxaPj(F)
Dim SrcFfnAy$()
   Dim S
   SrcFfnAy = AyWhLikAy(PthFfnAy(SrcPth), SslSy("*.bas *.cls"))
   For Each S In SrcFfnAy
       P.ImpSrcFfn S
   Next
PjRmvOptCmpDbLin P
PjImpRf P, SrcPth
PjSav P
End Sub

Sub XlsAddFxaNm(A As Excel.Application, FxaNm$)
Dim F$: F = FxaNmFxa(FxaNm)
If F = "" Then Exit Sub
A.AddIns.Add FxaNmFxa(FxaNm)
End Sub

Sub Z()
Z_FxaPj
End Sub

Private Sub Z_FxaPj()
Dim Fxa$
Fxa = TmpFxa
Ept = Fxa
GoSub Tst
Exit Sub
Tst:
    Act = FxaPj(Fxa).Filename
    C
    Return
End Sub

Private Sub Z_XlsAddFxaNm()
XlsAddFxaNm Excel.Application, "QIde0"
End Sub
