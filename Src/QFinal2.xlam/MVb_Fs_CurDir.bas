Attribute VB_Name = "MVb_Fs_CurDir"
Option Explicit
Function CurFnAy(Optional Spec$ = "*") As String()
CurFnAy = PthFnAy(CurDir, Spec)
End Function

Function CurSubFdrAy(Optional Spec$ = "*") As String()
CurSubFdrAy = PthSubFdrAy(CurDir)
End Function
