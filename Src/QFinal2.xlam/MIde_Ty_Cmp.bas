Attribute VB_Name = "MIde_Ty_Cmp"
Option Explicit
Function CvCmp(A) As VBComponent
Set CvCmp = A
End Function

Function CvCmpTyAy(CmpTyAy0$) As vbext_ComponentType()
Dim X, O() As vbext_ComponentType
For Each X In SslSy(CmpTyAy0)
    Push O, CmpStrTy(X)
Next
CvCmpTyAy = O
End Function


Function CmpStrTy(Sht) As vbext_ComponentType
Select Case Sht
Case "Doc": CmpStrTy = vbext_ComponentType.vbext_ct_Document
Case "Cls": CmpStrTy = vbext_ComponentType.vbext_ct_ClassModule
Case "Mod": CmpStrTy = vbext_ComponentType.vbext_ct_StdModule
Case "Frm": CmpStrTy = vbext_ComponentType.vbext_ct_MSForm
Case "ActX": CmpStrTy = vbext_ComponentType.vbext_ct_ActiveXDesigner
Case Else: Stop
End Select
End Function

Function CmpTyStr$(A As vbext_ComponentType)
Select Case A
Case vbext_ComponentType.vbext_ct_Document:    CmpTyStr = "Doc"
Case vbext_ComponentType.vbext_ct_ClassModule: CmpTyStr = "Cls"
Case vbext_ComponentType.vbext_ct_StdModule:   CmpTyStr = "Std"
Case vbext_ComponentType.vbext_ct_MSForm:      CmpTyStr = "Frm"
Case vbext_ComponentType.vbext_ct_ActiveXDesigner: CmpTyStr = "ActX"
Case Else: Stop
End Select
End Function
