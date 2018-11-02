Attribute VB_Name = "MVb_Fs_Ft"
Option Explicit

Sub FtBrw(A)
Shell "code.cmd """ & A & """", vbHide
'Shell "notepad.exe """ & A & """", vbMaximizedFocus
End Sub

Sub FtConstBrw(A, VarNm)
AyBrw FtConstLy(A, VarNm)
End Sub

Function FtConstLy(A, VarNm) As String()
'FtConstLy = LyConstLy(FtLy(A), VarNm)
End Function

Function FtDic(A) As Dictionary
Set FtDic = LyDic(FtLy(A))
End Function

Function FtEns(A)
If Not FfnIsExist(A) Then StrWrt "", A
FtEns = A
End Function

Sub FtIni(A)
If FfnIsExist(A) Then Exit Sub
StrWrt "", A
End Sub

Function FtInpHd%(A)
Dim O%
O = FreeFile(1)
Open A For Input As #O
FtInpHd = O
End Function

Function FtMayLines$(A)
On Error GoTo X
FtMayLines = FtLines(A)
Exit Function
X:
Debug.Print "FtMayLines: cannot read Ft[" & A & "] Err=[" & Err.Description & "]"
End Function
Function FtLines$(A)
If FfnSz(A) <= 0 Then Exit Function
FtLines = Fso.GetFile(A).OpenAsTextStream.ReadAll
End Function

Function FtLy(A) As String()
FtLy = SplitLines(FtLines(A))
End Function

Function FtAppFilNo%(A)
Dim O%: O = FreeFile(1)
Open A For Append As #O
FtAppFilNo = O
End Function

Function FtInpFilNo%(Ft)
Dim O%: O = FreeFile(1)
Open Ft For Input As #O
FtInpFilNo = O
End Function

Function FtOupFilNo%(A)
Dim O%: O = FreeFile(1)
Open A For Output As #O
FtOupFilNo = O
End Function

Sub FtRmvFst4Lines(Ft$)
Dim A$: A = Fso.GetFile(Ft).OpenAsTextStream.ReadAll
Dim B$: B = Left(A, 55)
Dim C$: C = Mid(A, 56)
Dim B1$: B1 = Replace("VERSION 1.0 CLASS|BEGIN|  MultiUse = -1  'True|END|", "|", vbCrLf)
If B <> B1 Then Stop
Fso.CreateTextFile(Ft, True).Write C
End Sub
