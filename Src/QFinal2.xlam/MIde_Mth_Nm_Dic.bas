Attribute VB_Name = "MIde_Mth_Nm_Dic"
Option Explicit

Function MdMthNmDic(A As CodeModule) As Dictionary
Set MdMthNmDic = SrcMthNmDic(MdSrc(A))
End Function

Function SrcMthNmDic(A$()) As Dictionary
'Return a dic with Key=MthNm and Val=MthLines with merging of Prp-Get-Let-Set of one-Lines
Dim Ix, MthNm$, Lines$
Dim O As New Dictionary
O.Add "*Dcl", SrcDclLines(A)
For Each Ix In AyNz(SrcMthIx(A))
    Lines = SrcMthIxLinesWithTopRmk(A, CLng(Ix))
    MthNm = LinMthNm(A(Ix)): If MthNm = "" Then Stop
    DicAddOrUpd O, MthNm, Lines, vbCrLf & vbCrLf
Next
Set SrcMthNmDic = O
End Function

Private Sub Z_SrcMthNmDic()
DicBrw SrcMthNmDic(CurSrc)
End Sub
