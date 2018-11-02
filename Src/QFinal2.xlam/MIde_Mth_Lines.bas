Attribute VB_Name = "MIde_Mth_Lines"
Option Explicit
Function MthDDNmLines$(MthDDNm$)
MthDDNmLines = MthLines(DDNmMth(MthDDNm))
End Function

'aaa
Private Property Get XX1()

End Property

'BB
Property Let XX1(V)

End Property

Sub Z_MthDDNmLines()
GoTo ZZ
ZZ:
Debug.Print MthDDNmLines("QIde.MIde_Mth_Lines.ZZ_MthDDNmLines")
End Sub

Function CurMthBdyLines$()
CurMthBdyLines = MdMthBdyLines(CurMd, CurMthNm$)
End Function

'aa
Private Sub Z_MthLines()
Debug.Print MthLines(Mth(CurMd, "XX1"), WithTopRmk:=True)
End Sub


Function MthEndLin$(MthLin$)
Dim A$
A = LinMthKd(MthLin): If A = "" Then Stop
MthEndLin = "End " & A
End Function

Function MdMthBdyLines$(A As CodeModule, MthNm$)
MdMthBdyLines = SrcMthBdyLines(MdBdyLy(A), MthNm)
End Function

Function MthBdyLy(A As CodeModule, MthNm$) As String()
MthBdyLy = SrcMthBdyLy(MdBdyLy(A), MthNm)
End Function

Function MthLines$(A As Mth, Optional WithTopRmk As Boolean)
MthLines = SrcMthLines(MdBdyLy(A.Md), A.Nm, WithTopRmk)
End Function

Function MthLinesWithTopRmk$(A As Mth)
MthLinesWithTopRmk = MthLines(A, WithTopRmk:=True)
End Function

Function MthLinCnt%(A As Mth)
MthLinCnt = FmCntAyLinCnt(MthFC(A))
End Function
