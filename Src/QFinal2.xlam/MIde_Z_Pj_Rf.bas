Attribute VB_Name = "MIde_Z_Pj_Rf"
Option Explicit

Function CurPjRfFfnAy() As String()
CurPjRfFfnAy = PjRfFfnAy(CurPj)
End Function

Function CurPjRfFmt() As String()
CurPjRfFmt = AyAlign2T(PjRfLy(CurPj))
End Function

Function CurPjRfLy() As String()
CurPjRfLy = PjRfLy(CurPj)
End Function
Sub CpyRf()

'Sub AcsCpyRf(A As Access.Application, Fb$)
'Dim R$()
'    R = PjRfFfnAy(A.Vbe.ActiveVBProject)
'Dim ToAcs As Access.Application
'Set ToAcs = FbAcs(Fb)
'PjAddRfFfnAy ToAcs.Vbe.ActiveVBProject, R
'AcsQuit ToAcs
'End Sub

End Sub

Function EmpRfAy() As Reference()
End Function

Function CvPjRf(A) As VBIDE.Reference
Set CvPjRf = A
End Function


Function RfNy() As String()
RfNy = CurPjRfNy
End Function

Function CurPjRfNy() As String()
CurPjRfNy = PjRfNy(CurPj)
End Function

Sub PjAddRf(A As VBProject, RfNm, PjFfn)
If PjHasRf(A, RfNm) Then Exit Sub
A.References.AddFromFile PjFfn
End Sub

Function PjRfAy(A As VBProject) As Reference()
PjRfAy = ItrAyInto(A.References, PjRfAy)
End Function

Sub PjRfBrw(A As VBProject)
AyBrw PjRfLy(A)
End Sub

Function PjRfCfgFfn$(A As VBProject)
PjRfCfgFfn = PjSrcPth(A) & "PjRf.Cfg"
End Function

Sub PjRfDmp(A As VBProject)
AyDmp PjRfLy(A)
End Sub

Function PjRfFfnAy(A As VBProject) As String()
PjRfFfnAy = ItrPrpSy(A.References, "FullPath")
End Function

Function PjRfLy(A As VBProject) As String()
Dim RfAy() As Reference
    RfAy = PjRfAy(A)
Dim O$()
Dim Ny$(): Ny = OyNy(RfAy)
Ny = AyAlignL(Ny)
Dim J%
For J = 0 To UB(Ny)
    Push O, Ny(J) & " " & RfFfn(RfAy(J))
Next
PjRfLy = O
End Function

Function PjRfNmRfFfn$(A As VBProject, RfNm$)
PjRfNmRfFfn = PjPth(A) & RfNm & ".xlam"
End Function

Function PjRfNy(A As VBProject) As String()
PjRfNy = ItrNy(A.References)
End Function

Function RfFfn$(A As Reference)
On Error Resume Next
RfFfn = A.FullPath
End Function

Function RfLin$(A As VBIDE.Reference)
RfLin = A.Name & " " & QuoteSqBkt(A.FullPath) & " " & QuoteSqBkt(A.Description)
End Function

Function RfPth$(A As VBIDE.Reference)
On Error Resume Next
RfPth = A.FullPath
End Function

Function RfToStr$(A As VBIDE.Reference)
With A
   RfToStr = .Name & " " & RfPth(A)
End With
End Function
