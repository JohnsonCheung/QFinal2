Attribute VB_Name = "MVb__Cv"
Option Explicit

Function CvPth$(ByVal A$)
If A = "" Then
    A = CurDir
End If
CvPth = PthEnsSfx(A)
End Function

Function CvSy(A) As String()
Select Case True
Case IsSy(A): CvSy = A
Case IsArray(A): CvSy = AySy(A)
Case Else: CvSy = ApSy(CStr(A))
End Select
End Function

Function CvAy(A) As Variant()
CvAy = A
End Function

Function CvFTIx(A) As FTIx
Set CvFTIx = A
End Function

Function CvFTNo(A) As FTNo
Set CvFTNo = A
End Function
Function CvNothing(A)
If IsEmpty(A) Then Set CvNothing = Nothing: Exit Function
Set CvNothing = A
End Function

Function CvNy(Ny0) As String()
Select Case True
Case IsMissing(Ny0)
Case IsStr(Ny0): CvNy = SslSy(Ny0)
Case IsSy(Ny0): CvNy = Ny0
Case IsArray(Ny0): CvNy = AySy(Ny0)
Case Else: Stop
End Select
End Function
