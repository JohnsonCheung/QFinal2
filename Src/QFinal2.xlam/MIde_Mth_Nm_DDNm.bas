Attribute VB_Name = "MIde_Mth_Nm_DDNm"
Option Explicit
Function LinMthDDNm$(A$)
'Return Nm.ShtKd.ShtMdy or Blank
Dim C$(): C = LinMthNmBrk(A): If Sz(C) = 0 Then Exit Function
LinMthDDNm = C(0) & "." & C(1) & "." & C(2)
End Function


Function MdMthDDNmDic(A As CodeModule) As Dictionary
Set MdMthDDNmDic = DicAddKeyPfx(SrcDic(MdSrc(A)), MdDNm(A) & ".")
End Function


Function SrcMthDDNmDic(A$()) As Dictionary
Dim Ix
Set SrcMthDDNmDic = New Dictionary
SrcMthDDNmDic.Add "*Dcl", SrcDclLines(A)
For Each Ix In AyNz(SrcMthIx(A))
    SrcMthDDNmDic.Add LinMthDDNm(A(Ix)), SrcMthIxLinesWithTopRmk(A, CLng(Ix))
Next
End Function


Function IsMthDDNmSel(A, B As WhMth) As Boolean
IsMthDDNmSel = MthNmBrkIsSel(SplitDot(A), B)
End Function

Function IsMthDNm(Nm) As Boolean
IsMthDNm = Sz(Split(Nm, ".")) = 3
End Function



Function MthNmBrkDDNm$(MthNmBrk$())
Select Case Sz(MthNmBrk)
Case 0:
Case 3: MthNmBrkDDNm = MthNmBrk(0) & "." & MthNmBrk(1) & "." & MthNmBrk(2)
Case Else: Stop
End Select
End Function
