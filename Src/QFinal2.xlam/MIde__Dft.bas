Attribute VB_Name = "MIde__Dft"
Option Explicit
Function DftMdByNm(MdNm$) As CodeModule
If MdNm = "" Then
    Set DftMdByNm = CurMd
Else
    Set DftMdByNm = Md(MdNm)
End If
End Function
Function DftMd(A As CodeModule) As CodeModule
If IsNothing(A) Then
   Set DftMd = CurMd
Else
   Set DftMd = A
End If
End Function

Function DftPj(A As VBProject) As VBProject
If IsNothing(A) Then
   Set DftPj = CurPj
Else
   Set DftPj = A
End If
End Function


Function DftMdyAy(A$) As String()
DftMdyAy = CvNy(A)
End Function

Function DftMth(MthDNm0$) As Mth
If MthDNm0 = "" Then
    Set DftMth = CurMth
    Exit Function
End If
Set DftMth = DDNmMth(MthDNm0)
End Function

Function DftPjByNm(PjNm$) As VBProject
If PjNm = "" Then
    Set DftPjByNm = CurPj
Else
    Set DftPjByNm = Pj(PjNm)
End If
End Function

Function DftFun(FunDNm0$) As Mth
If FunDNm0 = "" Then
    Dim M As Mth
    Set M = CurMth
    If IsMthFun(M) Then
        Set DftFun = M
    End If
Else
End If
Stop '
End Function
