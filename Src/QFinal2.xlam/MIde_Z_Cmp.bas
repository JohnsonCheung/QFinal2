Attribute VB_Name = "MIde_Z_Cmp"
Option Explicit
Function CurCmp() As VBComponent
Set CurCmp = CurMd.Parent
End Function

Sub CmpRmv(A As VBComponent)
A.Collection.Remove A
End Sub

Function CmpIsClsOrMod(A As VBComponent) As Boolean
Select Case A.Type
Case vbext_ct_ClassModule, vbext_ct_StdModule: CmpIsClsOrMod = True
End Select
End Function
