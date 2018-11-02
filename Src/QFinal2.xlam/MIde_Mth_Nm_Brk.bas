Attribute VB_Name = "MIde_Mth_Nm_Brk"
Option Explicit

Sub MthBrkAsg(A As Mth, OMdy$, OMthTy$)
Dim L$: L = MthDcl(A)
OMdy = TakMdy(L)
OMthTy = LinMthTy(L)
End Sub

Function MthBrkAyDDNy(A() As Variant) As String()
MthBrkAyDDNy = DryJnDotSy(A)
End Function

Function MthBrkAyWhDup(A()) As Variant()
'MthBrk is Sy of ShtMdy ShtTy Nm
Dim Dry(): Dry = DryWhColInAy(A, 0, Array("", "Pub")) '
MthBrkAyWhDup = DryWhColHasDup(Dry, 2)
End Function

Function MthNmBrkNm$(MthNmBrk$())
Select Case Sz(MthNmBrk)
Case 0:
Case 3: MthNmBrkNm = MthNmBrk(0)
Case Else: Stop
End Select
End Function

Function MthNmBrkIsSel(MthNmBrk$(), B As WhMth) As Boolean
Select Case Sz(MthNmBrk)
Case 0: Exit Function
Case 3: MthNmBrkIsSel = IsMthSel(MthNmBrk(0), MthNmBrk(1), MthNmBrk(2), B)
Case Else: Stop
End Select
End Function

Function IsMthSel(MthNm$, ShtTy$, ShtMdy$, A As WhMth) As Boolean
If MthNm = "" Then Exit Function
If IsNothing(A) Then IsMthSel = True: Exit Function
If Not AySel(A.InShtMdy, ShtMdy) Then Exit Function
If Not AySel(A.InShtKd, MthShtKd(ShtTy)) Then Exit Function
IsMthSel = IsNmSel(MthNm, A.Nm)
End Function

Function MthNmBrkAyWh(A() As Variant, B As WhMth) As Variant()
Dim Brk
For Each Brk In AyNz(A)
    If MthNmBrkIsSel(CvSy(Brk), B) Then PushI MthNmBrkAyWh, Brk
Next
End Function

Function LinMthNmBrk(A) As String()
LinMthNmBrk = ShfMthNmBrk(CStr(A))
End Function

Function MthNmBrkAyNy(A() As Variant) As String()
MthNmBrkAyNy = DryDistSy(A, 2)
End Function

Sub LinMthNmBrkAsg(A$, OShtMdy$, OShtTy$, ONm$)
Dim L$: L = A
OShtMdy = ShfShtMdy(L)
OShtTy = ShfMthShtTy(L)
ONm = TakNm(L)
End Sub
