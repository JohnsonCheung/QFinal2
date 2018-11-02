Attribute VB_Name = "MIde_Srt_Rpt"
Option Explicit
Private Type MdSrtRpt
    MdIdxDt As Dt
    RptDic As Dictionary ' K is Module Name, V is DicCmpRsltLy
End Type

Private Function SrcSrtRpt(A$()) As String()
Dim A1 As Dictionary
Dim B1 As Dictionary
Set A1 = SrcDic(A)
Set B1 = SrcDic(SrcSrtLy(A))
SrcSrtRpt = DicCmpFmt(A1, B1, "BefSrt", "AftSrt")
End Function

Private Function ZZ_Src() As String()
ZZ_Src = MdSrc(CurMd)
End Function

Private Sub ZZ_SrcSrtRpt()
Brw SrcSrtRpt(CurSrc)
End Sub

Function CurMdSrtRpt() As String()
CurMdSrtRpt = MdSrtRpt(CurMd)
End Function

Function PjSrtCmpRptWb(A As VBProject, Optional Vis As Boolean) As Workbook
Dim A1 As MdSrtRpt
A1 = PjMdSrtRpt(A)
Dim O As Workbook: Set O = DicWb(A1.RptDic)
Dim Ws As Worksheet
Set Ws = WbAddWs(O, "Md Idx", AtBeg:=True)
Dim Lo As ListObject: Set Lo = DtLo(A1.MdIdxDt, WsA1(Ws))
LoCol_LnkWs Lo, "Md"
If Vis Then WbVis O
Set PjSrtCmpRptWb = O
End Function

Function PjSrtRpt(A As VBProject) As String()
Dim O$(), I
For Each I In AyNz(PjMdAy(A))
    PushAy O, MdSrtRpt(CvMd(I))
Next
PjSrtRpt = O
End Function

Function PjSrtRptWb(A As VBProject, Optional Vis As Boolean) As Workbook
Set PjSrtRptWb = DicWb(LyDic(PjMdSrtRptDic(A)))
Stop '
Dim O As Workbook: ' Set O = DicWb(A1.RptDic)
Dim Ws As Worksheet
Set Ws = WbAddWs(O, "Md Idx")
'Dim Lo As ListObject: Set Lo = DtLo(A1.MdIdxDt, WsA1(Ws))
'LoCol_LnkWs Lo, "Md"
'If Vis Then WbVis O
'Set PjSrtRptWb = O
Stop '
End Function

Private Function PjMdSrtRpt(A As VBProject) As MdSrtRpt _
'SrtCmpDic is a LyDic with Key as MdNm and value is SrtCmpLy
Dim MdAy() As CodeModule: 'MdAy = PjMdAy(A)
Dim MdNy$(): MdNy = OyNy(MdAy)
Dim LyAy()
Dim IsSam$(), IsDif$(), Sam As Boolean
    Dim J%, R As DCRslt
    For J = 0 To UB(MdAy)
        Push LyAy, MdSrtRpt(MdAy(J))
'        Sam = DCRsltIsSam(R)
'        Push IsSam, IIf(Sam, "*Sam", "")
'        Push IsDif, IIf(Sam, "", "*Dif")
    Next
With PjMdSrtRpt
    Set .RptDic = AyabDic(MdNy, LyAy)
    .MdIdxDt = Dt("Md-Bef-Aft-Srt", "Md Sam Dif", AyZipAp(MdNy, IsSam, IsDif))
End With
End Function

Function PjMdSrtRptDic(A As VBProject) As Dictionary 'Return a dic of [MdNm,SrtCmpFmt]
'SrtCmpDic is a LyDic with Key as MdNm and value is SrtCmpLy
Dim I, O As New Dictionary, Md As CodeModule
    For Each I In AyNz(PjMdAy(A))
        Set Md = I
        O.Add MdNm(Md), MdSrtRpt(CvMd(Md))
    Next
Set PjMdSrtRptDic = O
End Function
Function MdSrtRpt(A As CodeModule) As String()
MdSrtRpt = SrcSrtRpt(MdSrc(A))
End Function

Function MdSrtDic(A As CodeModule) As Dictionary
Set MdSrtDic = DicAddKeyPfx(SrcSrtDic(MdSrc(A)), MdNm(A) & ".")
End Function

Sub MdSrtRptBrw(A As CodeModule)
Brw MdSrtRpt(A)
End Sub

